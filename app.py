from flask import Flask, render_template, request, redirect, url_for, session
from docx import Document
import random
import os
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
import openpyxl
from flask_session import Session

import imaplib
import email
from email.header import decode_header
import threading
import time

# Email credentials for monitoring (use an app password if 2FA is on)
EMAIL_USER = "milind.t-pharmacy@msubaroda.ac.in"
EMAIL_PASS = "cwelnregwtezxbil"
IMAP_SERVER = "imap.gmail.com"

# Server control flag
server_online = False


def check_email_commands():
    global server_online
    while True:
        try:
            mail = imaplib.IMAP4_SSL("imap.gmail.com")
            mail.login(EMAIL_USER, EMAIL_PASS)
            mail.select("inbox")

            # Search for latest emails
            result, data = mail.search(None, 'SUBJECT "START"')
            if result == 'OK':
                ids = data[0].split()
                if ids:
                    print("START email found.")
                    server_online = True
            else:
                print("No START emails found.")

            result, data = mail.search(None, 'SUBJECT "SHUTDOWN"')
            if result == 'OK':
                ids = data[0].split()
                if ids:
                    print("SHUTDOWN email found.")
                    server_online = False
            else:
                print("No SHUTDOWN emails found.")

            mail.logout()
        except Exception as e:
            print("Error in check_email_commands:", e)

        time.sleep(30)  # Check every 30 seconds

app = Flask(__name__)
# Start the background thread for email control
email_thread = threading.Thread(target=check_email_commands, daemon=True)
email_thread.start()

app.secret_key = 'gpat-secret-key'
app.secret_key = 'gpat-secret-key'

# Configure server-side session
app.config['SESSION_TYPE'] = 'filesystem'  # Store session data on the server
app.config['SESSION_PERMANENT'] = False
app.config['SESSION_FILE_DIR'] = './flask_session_data'  # Optional: set a custom folder

Session(app)  # Initialize the session extension


# In-memory user database (replace with actual DB in production)
users = {}

# Email verification code storage
verification_codes = {}
verification_expiry = {}

# Load MCQs from the .docx file
import pandas as pd


def load_mcqs_from_excel_columns(excel_path):
    df = pd.read_excel(excel_path, header=None)  # No headers assumed

    questions = []
    num_questions = len(df.iloc[:, 0].dropna())

    for i in range(num_questions):
        question = str(df.iloc[i, 0]).strip()  # Column A
        options = [
            str(df.iloc[i, 2]).strip(),  # Column C
            str(df.iloc[i, 3]).strip(),  # Column D
            str(df.iloc[i, 4]).strip(),  # Column E
            str(df.iloc[i, 5]).strip()   # Column F
        ]
        answer = str(df.iloc[i, 6]).strip()  # Column G

        questions.append({
            'question': question,
            'options': options,
            'answer': answer,
            'status': 'Not Visited',
            'selected': None,
            'marked_for_review': False
        })

    return questions

# Example usage
questions = load_mcqs_from_excel_columns("FINAL GPAT EXCEL FILE_new.xlsx")



# Helper function to send email
def send_verification_email(to_email, code):
    sender_email = "milind.t-pharmacy@msubaroda.ac.in"
    sender_password = "cwelnregwtezxbil"
    subject = "GPAT Registration Verification Code"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = to_email
    msg['Subject'] = subject

    body = f"Your verification code is: {code}\n\nThis code will expire in 10 minutes."
    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, to_email, msg.as_string())
        server.quit()
    except Exception as e:
        print(f"Email sending failed: {e}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        name = request.form['name']
        email = request.form['email']
        password = request.form['password']
        code = request.form.get('code')

        if email in users:
            return "User already exists."

        if not code:
            verification_code = str(random.randint(100000, 999999))
            verification_codes[email] = verification_code
            verification_expiry[email] = datetime.now() + timedelta(minutes=10)
            send_verification_email(email, verification_code)
            return render_template('register.html', message='Verification code sent to your email.', name=name, email=email, password=password)
        else:
            if email not in verification_codes or datetime.now() > verification_expiry.get(email, datetime.min):
                return "Verification code expired or not found. Please register again."

            if verification_codes.get(email) == code:
                users[email] = password
                df = pd.DataFrame([[name, email, password]], columns=['Name', 'Email', 'Password'])
                if os.path.exists('users.xlsx'):
                    existing = pd.read_excel('users.xlsx')
                    df = pd.concat([existing, df], ignore_index=True)
                df.to_excel('users.xlsx', index=False)
                return redirect(url_for('login'))
            else:
                return "Invalid verification code."

    return render_template('register.html')
@app.before_request
def check_server_status():
    allowed_paths = ['/offline', '/static', '/favicon.ico']
    if not server_online:
        # Let only static resources and offline page through
        if not any(request.path.startswith(path) for path in allowed_paths):
            return redirect('/offline')

@app.route('/offline')
def offline_page():
    return "<h1>🔒 Server is currently offline. Please try again later.</h1>"

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        userid = request.form['userid']
        password = request.form['password']
        if users.get(userid) == password:
            # Try to load user's name from Excel
            df = pd.read_excel('users.xlsx')
            user_row = df[df['Email'] == userid]
            if not user_row.empty:
                user_name = user_row.iloc[0]['Name']
                session['user'] = {
                    'name': user_name,
                    'email': userid
                }
            else:
                session['user'] = {
                    'name': 'Unknown',
                    'email': userid
                }
            return redirect(url_for('start_test'))
        return "Invalid credentials."
    return render_template('login.html')




@app.route('/start_test', methods=['GET','POST'])
def start_test():
    if 'user' not in session:
        return redirect(url_for('login'))

    session['questions'] = load_mcqs_from_excel_columns('FINAL GPAT EXCEL FILE_new.xlsx')
    session['q_index'] = 0
    session['start_time'] = datetime.now().timestamp()
    session['time_left'] = 180 * 60  # 3 hours
    return redirect(url_for('quiz'))


# ✅ Quiz display and navigation
@app.route('/quiz', methods=['GET', 'POST'])
def quiz():
    if 'questions' not in session or 'user' not in session:
        
        return redirect(url_for('start_test'))

    if request.method == 'POST':
        if 'submit' in request.form:
            elapsed = datetime.now().timestamp() - session.get('start_time', 0)
            if elapsed < 180 * 60:
                return "You can only submit after the timer ends."

        selected = request.form.get('option')
        mark_review = 'mark_review' in request.form
        index = session['q_index']
        questions = session['questions']

        # Update current question state
        questions[index]['selected'] = selected
        questions[index]['marked_for_review'] = mark_review
        questions[index]['status'] = (
            'Answered-Marked for Review' if mark_review and selected else
            'Marked for Review' if mark_review else
            'Answered' if selected else 'Not Answered'
        )
        session['questions'] = questions

        # Navigation
        if 'next' in request.form:
            session['q_index'] = min(index + 1, len(questions) - 1)
        elif 'prev' in request.form:
            session['q_index'] = max(index - 1, 0)
        elif 'palette_nav' in request.form:
            session['q_index'] = int(request.form.get('palette_nav'))

    index = session['q_index']
    question = session['questions'][index]
    palette = [(i + 1, q['status']) for i, q in enumerate(session['questions'])]
    user = session.get('user',{})
    
    return render_template('exam_ui_gpat.html',
                           qn=index + 1,
                           question=question,
                           total=len(session['questions']),
                           palette=palette,
                           time_left=session['time_left'],
                           user=user)
import pandas as pd

def send_result_email(to_email, name, score, max_score, attempted, correct_answers, incorrect_answers, unattempted):
    sender_email = "milind.t-pharmacy@msubaroda.ac.in"
    sender_password = "cwelnregwtezxbil"
    subject = "Your GPAT Mock Test Result"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = to_email
    msg['Subject'] = subject

    body = f"""Dear {name},

Thank you for completing the GPAT Mock Test.

Here is your result summary:

- Total Score: {score} / {max_score}
- Attempted Questions: {attempted}
- Correct Answers: {correct_answers}
- Incorrect Answers: {incorrect_answers}
- Unattempted Questions: {unattempted}

All the best for your preparation!

Regards,
GPAT Mock Test Team
"""

    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, to_email, msg.as_string())
        server.quit()
        print("Result email sent successfully.")
    except Exception as e:
        print(f"Email sending failed: {e}")


@app.route('/result')
def result():
    if  session.get('show_result'):
        return redirect(url_for('thank_you'))
    score = 0
    attempted = 0
    unattempted = 0
    correct_answers = 0
    incorrect_answers = 0

    # Prepare the list to store the responses
    responses = []

    for q in session['questions']:
        if q['selected'] is not None:
            attempted += 1

            # Compare the first letter of selected with last letter of answer
            selected_letter = str(q['selected']).strip()[0].upper()
            correct_letter = str(q['answer']).strip()[-1].upper()

            if selected_letter == correct_letter:
                correct_answers += 1
                score += 4
            else:
                incorrect_answers += 1
                score -= 1
        else:
            unattempted += 1



        # Save the response data for each question
        responses.append({
            'User': session['user']['email'],
            'Question': q['question'],
            'Selected Answer': q['selected'],
            'Correct Answer': q['answer'],
            'Is Correct': q['selected'] == q['answer'] if q['selected'] else False
        })

    # Save all responses to Excel
    responses_df = pd.DataFrame(responses)

    # Append to existing file if it exists
    if os.path.exists('responses.xlsx'):
        existing_data = pd.read_excel('responses.xlsx')
        responses_df = pd.concat([existing_data, responses_df], ignore_index=True)

    responses_df.to_excel('responses.xlsx', index=False)

    # The test is out of 500 marks, so the total score is based on number of questions × 4
    max_score = len(session['questions']) * 4

    user_email = session['user']['email']
    user_name = session['user'].get('name', 'Candidate')

    # Send result via email
    send_result_email(user_email, user_name, score, max_score, attempted, correct_answers, incorrect_answers, unattempted)
    # Store the result data in session temporarily
    session['score_data'] = {
    'score': score,
    'total': len(session['questions']),
    'attempted': attempted,
    'correct_answers': correct_answers,
    'incorrect_answers': incorrect_answers,
    'unattempted': unattempted,
    'max_score': max_score
     }
    session['show_result'] = True
    

    return render_template('result.html',
                           score=score,
                           total=len(session['questions']),
                           attempted=attempted,
                           correct_answers=correct_answers,
                           incorrect_answers=incorrect_answers,
                           unattempted=unattempted,
                           max_score=max_score)


@app.route('/update_time', methods=['POST'])
def update_time():
    data = request.get_json()
    session['time_left'] = data['time_left']  # Update time_left in the session
    return '', 204  # No content response
@app.route('/thank_you')
def thank_you():
    session.clear()  # ensure no session data remains
    return render_template('thank_you.html')

@app.after_request
def add_header(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "-1"
    return response
# @app.route('/tab-switch', methods=['POST'])
# def tab_switch():
#     session.clear()  # Optional: clear exam session
#     return '', 204
# @app.route('/exit')
# def exit_exam():
#     return render_template('exam_terminated.html')

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host='0.0.0.0', port=port)
    

