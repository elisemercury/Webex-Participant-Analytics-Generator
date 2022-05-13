from flask import Flask, redirect, request, render_template, send_file, session
from flask_session import Session
import requests
import urllib.parse
import datetime
import xlsxwriter
import os
import pytz
import redis
from redis import Redis

# --- base variables
base_URL = os.environ["BASE_URL"]
myClientID = os.environ["CLIENT_ID"] 
myClientSecret = os.environ["CLIENT_SECRET"]
myScope = os.environ["SCOPE"]
myRedirectURI = os.environ["REDIRECT_URL"]
APP_URL = os.environ["APP_URL"]
datastore = os.environ['REDIS_URL']


app = Flask(__name__)
SESSION_TYPE = 'redis'
SESSION_REDIS = Redis.from_url(url=datastore)
app.config.from_object(__name__)
Session(app)

# --- landing page
@app.route('/')
def hello():
    login_msg = "Please log in with Webex to start."
    return render_template('login.html', app_url=APP_URL, login_msg=login_msg)

# --- perform login
@app.route('/gologin')
def gologin():
    try:
        wx_domain = request.form['wxDomain'].split()
    except:
        print("isuuueee")
    query = request.url

    if 'code' not in query:
        # Create the redirect URL
        oauthRedirectUrl = get_oauthRedirectUrl(myClientID, myRedirectURI, myScope)
        # Redirect user to this URL
        return redirect(oauthRedirectUrl)

    if 'code' in query:
        # Extract the 'code' from the URL
        teamsAuthCode = query.split('=', 1)[-1]

        # With the 'code', now get your real accesss token.
        session["myAccessToken"] = get_token(myRedirectURI, teamsAuthCode, myClientID, myClientSecret)
        print(session["myAccessToken"])
        

    if session["myAccessToken"]:
        session["myUsername"] = get_myDetails(session["myAccessToken"])
        return redirect('/main')
    else:
        login_msg = "There was an issue authenticating you with Webex. Please try again."
        return render_template('login.html', app_url=APP_URL, login_msg=login_msg)

# create redirect URL
def get_oauthRedirectUrl(myClientID, myRedirectURI, myScope):
	oauthRedirectUrl = base_URL+"v1/authorize"
	oauthRedirectUrl += "?response_type=code"
	oauthRedirectUrl += "&client_id=" + myClientID
	oauthRedirectUrl += "&redirect_uri=" + str(urllib.parse.quote(myRedirectURI, safe='~@#$&()*!+=;,.?\''))
	oauthRedirectUrl += "&scope=" + str(urllib.parse.quote(myScope, safe='_'))
	return oauthRedirectUrl

# get the user's Webex access token
def get_token(myRedirectURI, teamsAuthCode, myClientID, myClientSecret):
    data = {'grant_type': 'authorization_code', 'redirect_uri': myRedirectURI, 'code': teamsAuthCode, 'client_id': myClientID, 'client_secret': myClientSecret}
    header = {'content-type': 'application/x-www-form-urlencoded'}
    myAccessToken = False
    try:
        req = requests.post(base_URL+"v1/access_token", headers=header, data=data)
        response = req.json()
        myAccessToken = response['access_token']
    except:
        myAccessToken = False
    return myAccessToken

# get Webex account information
def get_myDetails(mytoken):
    header = {'Authorization': "Bearer " + mytoken,'content-type': 'application/json; charset=utf-8'}
    result = requests.get(url=base_URL+"v1/people/me", headers=header)
    return result.json()["userName"]

# login successfull, now enter meeting nr to fetch data for
@app.route('/main')
def home():
    try:
        if session["myUsername"]:
            return render_template('main-fetch.html', app_url=APP_URL, username=session["myUsername"])
    except:
        return redirect('/logout')

# --- fetch participant data by meeting number
@app.route('/main', methods=['POST'])
def post_meeting_nr():
    meeting_nr = "".join(request.form['meeting_nr'].split())
    # fetch meeting_id by meeting number
    try:
        meeting_id, session["meeting_name"], session["meeting_date"] = get_meetingID(session["myAccessToken"], meeting_nr)      
        
        # fetch participant data for meeting id
        participant_info = get_participant_info(session["myAccessToken"], meeting_id)
        if not participant_info:
            notification = "Could not fetch participant data for meeting number: " + meeting_nr
            return render_template('main-fetch.html', app_url=APP_URL, username=session["myUsername"], notification=notification)            
        export = create_xlsx_report(participant_info)
        if not export:
            notification = "Could not create participant report for meeting number: " + meeting_nr
            return render_template('main-fetch.html', app_url=APP_URL, username=session["myUsername"], notification=notification)
        
        session["meeting_nr_formatted"] = meeting_nr[0:4] + " " + meeting_nr[4:7] + " " + meeting_nr[7:]
        return redirect('/success')
    except Exception as e:
        print(e)
        try:
            if session["myUsername"]:
                notification = "Could not fetch meeting data for meeting number: " + meeting_nr
                return render_template('main-fetch.html', app_url=APP_URL, username=session["myUsername"], notification=notification)
        except:
            return redirect('/logout')

# meeting ID from meeting Nr
def get_meetingID(mytoken, meeting_nr):
    url = f'{base_URL}v1/meetings?meetingNumber={meeting_nr}'
    data = {"meetingNumber": meeting_nr}
    header = {'Authorization': "Bearer " + mytoken,'content-type': 'application/json; charset=utf-8'}
    result = requests.get(url, json=data, headers=header)
    if result.status_code != 200:
        return False
    else:
        meeting_id = result.json()['items'][0]['id']
        meeting_name = result.json()['items'][0]['title']
        meeting_date = result.json()['items'][0]['start'][0:10]
        return meeting_id, meeting_name, meeting_date

# participant data from meeting ID
def get_participant_info(mytoken, meeting_id):
    url = f'{base_URL}v1/meetingParticipants?meetingId={meeting_id}'
    data = {"meetingId": meeting_id}
    header = {'Authorization': "Bearer " + mytoken,'content-type': 'application/json; charset=utf-8'}
    result = requests.get(url, json=data, headers=header)
    if result.status_code != 200:
        return False
    else:
        return result.json()['items']

# create participant report
def create_xlsx_report(particpant_info):
    try:
        workbook = xlsxwriter.Workbook(session["meeting_name"] + "_" + session["meeting_date"] + '_participant_analytics.xlsx')
        worksheet = workbook.add_worksheet()

        bold = workbook.add_format({'bold': True})
        date_format = 'hh:mm:ss'
        date_format = workbook.add_format({'num_format': date_format, "align": "left"})

        worksheet.set_column(0, 0, 30)
        worksheet.set_column(1, 1, 25)
        worksheet.set_column(2, 4, 15)
        worksheet.set_column(5, 5, 20)

        worksheet.write(0, 0, "Participant Name", bold)
        worksheet.write(0, 1, "Email", bold)
        worksheet.write(0, 2, "Joined Time", bold)
        worksheet.write(0, 3, "Left Time", bold)
        worksheet.write(0, 4, "Timezone", bold)
        worksheet.write(0, 5, "Total Attendence", bold)

        row, col = 1, 0

        occurences = list()
        for participant in particpant_info:
            occurences.append((datetime.datetime.strptime(participant["joinedTime"], '%Y-%m-%dT%H:%M:%SZ').date()))
        print(1)
        if len(set(occurences)) != 1:
            now = datetime.datetime.now(pytz.utc).date()
            recent_date = max(dt for dt in occurences if dt < now)
            print(2)
        else:
            recent_date = occurences[0]
            print(3)

        for participant in particpant_info:
            joined = datetime.datetime.strptime(participant["joinedTime"], '%Y-%m-%dT%H:%M:%SZ')
            if participant["email"][0:7] == "machine" and participant["devices"][0]["deviceType"] == "tp_endpoint":
                pass
            elif joined.date() == recent_date:
                print(4)
                worksheet.write(row, col, participant["displayName"])
                worksheet.write(row, col+1, participant["email"])

                left = datetime.datetime.strptime(participant["leftTime"], '%Y-%m-%dT%H:%M:%SZ')
                timezone = pytz.utc.localize(joined)
                print(5)
                worksheet.write_datetime(row, col+2, joined.time(), date_format)
                worksheet.write_datetime(row, col+3, left.time(), date_format)
                worksheet.write(row, col+4, str(timezone.tzname()) + " " + str(timezone)[19:])

                total = left - joined 
                print(6)
                worksheet.write_datetime(row, col+5, total, date_format)
                row += 1
            
        workbook.close()
        return True
    except Exception as e:
        print(e)
        return False

# --- successfully fetched participant data
@app.route("/success")
def success():
    try:
        if session["myUsername"]:
             return render_template('main-fetch-success.html', app_url=APP_URL, username=session["myUsername"], meeting_nr=session["meeting_nr_formatted"], meeting_name=session["meeting_name"])
    except:
        return redirect('/logout')

# --- download participant report
@app.route("/success", methods=['POST'])
def download_report():
    chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    for char in chars:
        session["meeting_name"] = session["meeting_name"].replace(char, "_")

    return send_file(session["meeting_name"] + "_" + session["meeting_date"] + '_participant_analytics.xlsx',
                     mimetype='application/vnd.ms-excel',
                     attachment_filename=session["meeting_name"] + "_" + session["meeting_date"] + '_participant_analytics.xlsx',
                     as_attachment=True)

# --- help page
@app.route("/help")
def help():
    return render_template('help.html', app_url=APP_URL)

# --- help page
@app.route("/logout")
def logout():
    for key in list(session.keys()):
        session.pop(key)
    return redirect('/')

# --- handle 404 errors
@app.errorhandler(404)
def page_not_found(e):
    login_msg = "Please log in with Webex to start."
    return render_template('login.html', app_url=APP_URL, login_msg=login_msg), 404

# --- run the app
if __name__ == '__main__':
    app.run()