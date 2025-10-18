from flask import Flask, render_template, request, session, redirect
from firebase import Firebase
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)
app.secret_key = "your_secret_here"

# Firebase config
firebaseConfig = {
    "apiKey": "AIzaSyBYc1L6EQ0Kao1BbmBxRG66PNyjX5udI94",
    "authDomain": "employee-portal-85b28.firebaseapp.com",
    "projectId": "employee-portal-85b28",
    "storageBucket": "employee-portal-85b28.appspot.com",
    "messagingSenderId": "923148564575",
    "appId": "1:923148564575:web:892d7eba7c9c66387e43ea",
    "databaseURL": "https://employee-portal-85b28-default-rtdb.firebaseio.com/"
}
firebase = Firebase(firebaseConfig)
auth = firebase.auth()

# Google Sheets config
SERVICE_ACCOUNT_FILE = 'service_account.json'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
gc = gspread.authorize(credentials)
workbook = gc.open_by_key('1AIm4sXqlcIeMVNwI8S_cJnxgJYhErjDSvxmlxbqtXr8')

def get_profile(email):
    try:
        sheet = workbook.worksheet('PROFILE')
        rows = sheet.get_all_records()
        for row in rows:
            if str(row.get('EMAIL', '')).lower() == str(email).lower():
                return row
    except Exception:
        pass
    return {}

def field_value(request, name):
    val = request.form.get(name, '').strip()
    return val if val else '0'


@app.route('/')
def home():
    return redirect('/login')


@app.route('/login', methods=['GET', 'POST'])
def login():
    message = ""
    if request.method == 'POST':
        email = request.form['email']
        password = request.form['password']
        try:
            user = auth.sign_in_with_email_and_password(email, password)
            session['user'] = user['localId']
            session['email'] = email
            return redirect('/dashboard')
        except Exception as e:
            message = f"Login failed: {str(e)}"
    return render_template('login.html', message=message)


@app.route('/dashboard')
def dashboard():
    if 'user' not in session or 'email' not in session:
        return redirect('/login')
    profile = get_profile(session['email'])
    return render_template('dashboard.html',
        name = profile.get('NAME', 'User'),
        email = profile.get('EMAIL', ''),
        facility = profile.get('FACILITY', ''),
        role = profile.get('ROLE', ''),
        sex = profile.get('Sex', ''),
        birthdate = profile.get('Birthdate', ''),
        age = profile.get('Age', ''),
        address = profile.get('Address', ''),
        contact_number = profile.get('Contact Number', ''),
        tasks = profile.get('TASKS', '')
    )


@app.route('/encode', methods=['GET', 'POST'])
def encode():
    if 'user' not in session or 'email' not in session:
        return redirect('/login')

    profile = get_profile(session['email'])
    name = profile.get('NAME', 'User')
    facility = profile.get('FACILITY', 'Not Set')
    role = profile.get('ROLE', 'Not Set').strip().upper()
    message = ""
    months = [
        "October", "November", "December", "January", "February", "March",
        "April", "May", "June", "July", "August", "September"
    ]

    # ---- Get targets from TAR ----
    tar_sheet = workbook.worksheet('TAR')
    tar_data = tar_sheet.get_all_values()

    targets = {
        "all_tested": 0,
        "sns_tested": 0,
        "sns_positive": 0,
        "index_tested": 0,
        "index_positive": 0,
        "all_positive": 0,
        "art_enrolled": 0,
        "prep_enrolled": 0
    }

    for row in tar_data:
        if len(row) >= 17 and str(row[0]).strip().upper() == role:
            def safe_int(val):
                try:
                    return int(val)
                except:
                    return 0
            targets = {
                "all_tested": safe_int(row[2]),
                "sns_tested": safe_int(row[4]),
                "sns_positive": safe_int(row[6]),
                "index_tested": safe_int(row[8]),
                "index_positive": safe_int(row[10]),
                "all_positive": safe_int(row[12]),
                "art_enrolled": safe_int(row[14]),
                "prep_enrolled": safe_int(row[16])
            }
            break


    # --- Handle form submission ---
    if request.method == 'POST':
        try:
            month = field_value(request, 'month')
            staff = name

            def fv(name):
                return field_value(request, name)

            sns_tested = [fv(n) for n in [
                'sns_tested_m','sns_tested_f','sns_tested_msm','sns_tested_tgw','sns_tested_total','sns_tested_rate']]
            sns_positive = [fv(n) for n in [
                'sns_positive_m','sns_positive_f','sns_positive_msm','sns_positive_tgw','sns_positive_total','sns_positive_rate']]
            index_tested = [fv(n) for n in [
                'index_tested_m','index_tested_f','index_tested_msm','index_tested_tgw','index_tested_total','index_tested_rate']]
            index_positive = [fv(n) for n in [
                'index_positive_m','index_positive_f','index_positive_msm','index_positive_tgw','index_positive_total','index_positive_rate']]
            all_tested = [fv(n) for n in [
                'all_tested_m','all_tested_f','all_tested_msm','all_tested_tgw','all_tested_total','all_tested_rate']]
            all_positive = [fv(n) for n in [
                'all_positive_m','all_positive_f','all_positive_msm','all_positive_tgw','all_positive_total','all_positive_rate']]
            art_enrolled = [fv(n) for n in [
                'art_enrolled_m','art_enrolled_f','art_enrolled_msm','art_enrolled_tgw','art_enrolled_total','art_enrolled_rate']]
            prep_enrolled = [fv(n) for n in [
                'prep_enrolled_m','prep_enrolled_f','prep_enrolled_msm','prep_enrolled_tgw','prep_enrolled_total','prep_enrolled_rate']]
            variance = [fv(n) for n in [
                'variance_sns','variance_index','variance_all','variance_art','variance_prep']]

            data_row = (
                [month, facility, staff, role]
                + sns_tested + sns_positive
                + index_tested + index_positive
                + all_tested + all_positive
                + art_enrolled + prep_enrolled + variance
            )

            ws = workbook.worksheet('Accomplishment')
            all_rows = ws.get_all_values()[3:]

            existing_row_idx = None
            for idx, row in enumerate(all_rows, start=4):
                if row and str(row[0]).strip().lower() == month.strip().lower() \
                   and str(row[1]).strip().lower() == facility.strip().lower():
                    existing_row_idx = idx
                    break

            if existing_row_idx:
                ws.update(f"A{existing_row_idx}:BN{existing_row_idx}", [data_row])
                message = f"✅ Accomplishment for {month} updated!"
            else:
                ws.append_row(data_row, table_range="A4")
                message = f"✅ Accomplishment for {month} saved!"

        except Exception as e:
            message = f"❌ An error occurred while saving: {str(e)}"

    return render_template(
        'encode.html',
        name=name,
        facility=facility,
        role=role,
        months=months,
        message=message,
        targets=targets
    )
from flask import jsonify

@app.route('/get_accomplishment', methods=['POST'])
def get_accomplishment():
    try:
        data = request.get_json()
        month = data.get('month', '').strip().lower()
        facility = data.get('facility', '').strip().lower()
        name = data.get('name', '').strip().lower()
        role = data.get('role', '').strip().lower()

        ws = workbook.worksheet('Accomplishment')
        all_rows = ws.get_all_values()[3:]  # skip header rows

        # column mapping
        name_map = [
            "sns_tested_m","sns_tested_f","sns_tested_msm","sns_tested_tgw","sns_tested_total","sns_tested_rate",
            "sns_positive_m","sns_positive_f","sns_positive_msm","sns_positive_tgw","sns_positive_total","sns_positive_rate",
            "index_tested_m","index_tested_f","index_tested_msm","index_tested_tgw","index_tested_total","index_tested_rate",
            "index_positive_m","index_positive_f","index_positive_msm","index_positive_tgw","index_positive_total","index_positive_rate",
            "all_tested_m","all_tested_f","all_tested_msm","all_tested_tgw","all_tested_total","all_tested_rate",
            "all_positive_m","all_positive_f","all_positive_msm","all_positive_tgw","all_positive_total","all_positive_rate",
            "art_enrolled_m","art_enrolled_f","art_enrolled_msm","art_enrolled_tgw","art_enrolled_total","art_enrolled_rate",
            "prep_enrolled_m","prep_enrolled_f","prep_enrolled_msm","prep_enrolled_tgw","prep_enrolled_total","prep_enrolled_rate",
            "variance_sns","variance_index","variance_all","variance_art","variance_prep"
        ]

        # find the correct row
        for row in all_rows:
            if len(row) < 4:
                continue
            if (
                str(row[0]).strip().lower() == month
                and str(row[1]).strip().lower() == facility
                and str(row[2]).strip().lower() == name
                and str(row[3]).strip().lower() == role
            ):
                record = {name_map[i]: val for i, val in enumerate(row[4:4 + len(name_map)])}
                return jsonify({"success": True, "record": record})

        return jsonify({"success": False, "message": "No record found for this month."})

    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')


if __name__ == "__main__":
    app.run(debug=True)
