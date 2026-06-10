
import streamlit as st
import datetime
import re
import platform
import uuid
import urllib.parse
import base64

OUTLOOK_AVAILABLE = False
try:
    if platform.system() == "Windows":
        import win32com.client
        import pythoncom
        OUTLOOK_AVAILABLE = True
except ImportError:
    pass

st.set_page_config(
    page_title="SPS WE-TD Training Invitation Template Generator",
    page_icon="mail",
    layout="wide"
)

st.markdown("""
<style>
    .main-header { font-size: 2.2rem; font-weight: bold; color: #1a73e8; margin-bottom: 0.5rem; }
    .sub-header { font-size: 1.1rem; color: #555; margin-bottom: 2rem; }
    .success-box { padding: 1rem; border-radius: 0.5rem; background-color: #d4edda; border: 1px solid #c3e6cb; color: #155724; }
    .warning-box { padding: 1rem; border-radius: 0.5rem; background-color: #fff3cd; border: 1px solid #ffc107; color: #856404; }
    .info-box { padding: 1rem; border-radius: 0.5rem; background-color: #cce5ff; border: 1px solid #b8daff; color: #004085; }
    .step-box { padding: 1rem; border-radius: 0.5rem; background-color: #f0f8ff; border: 1px solid #b8daff; margin: 10px 0; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">SPS WE-TD Training Invitation Template Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Auto-generate and send training calendar invites</div>', unsafe_allow_html=True)

if OUTLOOK_AVAILABLE:
    st.success("**Mode: Local (Outlook Direct Send)** - Calendar invites will be sent directly from your Outlook")
else:
    st.info("**Mode: Cloud (Calendar Invite)** - Generate a calendar invite file to send via your Outlook")

st.divider()

def process_attendees(attendees_text):
    if not attendees_text.strip():
        return []
    emails = []
    for entry in attendees_text.split(";"):
        entry = entry.strip()
        if not entry:
            continue
        if "@" in entry:
            emails.append(entry)
        else:
            emails.append(f"{entry}@amazon.com")
    return emails

col1, col2 = st.columns(2)

with col1:
    st.subheader("Training Details")
    skill_name = st.text_input("Skill / Training Name *", placeholder="e.g., NH MKT Training")
    batch_id = st.text_input("Batch ID *", placeholder="e.g., SPSSELLER-LNH-IND-BLR-2024-04-22-X-MKT33")
    skill_details = st.text_area("Skill Details (Topic Description) *", placeholder="e.g., New Hire Marketing Training - Module 3", height=80)
    trainer_name = st.text_input("Trainer Name *", placeholder="e.g., John Doe", help="Also used as email signature")
    shift_details = st.text_input("Shift Details and WOs *", placeholder="e.g., Shift A (9:00 AM - 6:00 PM IST), WO: 12345")
    adobe_connect_link = st.text_input("Adobe Connect Room Link *", placeholder="e.g., https://amazon.adobeconnect.com/room-name/")

with col2:
    st.subheader("Schedule and Recipients")
    col_date1, col_date2 = st.columns(2)
    with col_date1:
        start_date = st.date_input("Start Date *", value=datetime.date.today() + datetime.timedelta(days=7))
    with col_date2:
        end_date = st.date_input("End Date *", value=datetime.date.today() + datetime.timedelta(days=37))
    required_attendees = st.text_area("Required Attendees (Login IDs) *", placeholder="Enter login IDs separated by semicolons (;)\ne.g., johndoe; janedoe; smithj", help="@amazon.com added automatically", height=100)
    optional_attendees = st.text_area("Optional Attendees (Supervisors/STMs)", placeholder="Enter login IDs separated by semicolons (;)\ne.g., manager1; stm1", help="@amazon.com added automatically. griffin-btr@share.corp.amazon.com is always included.", height=80)
    sender_email = st.text_input("Your Email (Sender/Organizer) *", placeholder="e.g., kiksh@amazon.com")

st.divider()

def format_date_display(date_obj):
    day = date_obj.day
    if 4 <= day <= 20 or 24 <= day <= 30:
        suffix = "th"
    else:
        suffix = ["st", "nd", "rd"][day % 10 - 1]
    return date_obj.strftime(f"%B {day}{suffix}")

def generate_subject(skill, batch, start, end):
    duration = f"{format_date_display(start)} to {format_date_display(end)}"
    return f"{skill} | {batch} | {duration}"

def generate_html_body(skill_details, trainer_name, shift_details, adobe_link):
    return f"""<html><body style="font-family:'Segoe UI',Arial,sans-serif;font-size:14px;color:#333;line-height:1.6;"><div style="max-width:700px;margin:0 auto;padding:20px;"><p style="font-size:15px;">Hello Everyone!</p><p>You have been invited to join a training session on <strong>{skill_details}</strong></p><p>Your Trainer will be: <strong>{trainer_name}</strong>. Shift: <strong>{shift_details}</strong></p><p><strong>Please respond to this e-invite to confirm your attendance.</strong></p><hr style="border:none;border-top:1px solid #ddd;margin:20px 0;"><p><strong>Link to Adobe Connect Room</strong> (to be opened via the Adobe Connect App only)</p><p><a href="{adobe_link}" style="color:#1a73e8;font-weight:bold;">{adobe_link}</a></p><div style="background-color:#fff3cd;border:1px solid #ffc107;border-radius:5px;padding:12px;margin:15px 0;"><p style="margin:0;"><strong>WARNING - Note:</strong> Since all WW SPS-WE trainings may involve customer accounts, Slack is not an approved tool for sharing customer data. Please make sure Adobe Connect is installed and ready before your training begins.</p><p style="margin:8px 0 0 0;font-style:italic;">Supervisors and Managers listed as optional attendees are added for visibility only. If this training is not relevant to you, please skip it and/or decline the invite so it does not block your calendar.</p></div><hr style="border:none;border-top:1px solid #ddd;margin:20px 0;"><p><strong>First time using Adobe Connect?</strong></p><p>Please ensure that the Adobe Connect application is installed on your laptop. If it is not already installed, you can do so by following these steps:</p><p style="background-color:#e8f4fd;padding:10px;border-radius:5px;border-left:4px solid #1a73e8;"><strong>Start Menu &gt; Search for "Software Center" &gt; Open Software Center &gt; Search for "Adobe Connect" &gt; Click "Install"</strong></p><p>Once installed, please copy the room link, paste it in the application and click on <strong>Continue</strong>. Now select <strong>Guest</strong> as the option and type your full name. Post that, click on <strong>Enter Room</strong>. Please join <strong>5-10 minutes early</strong> to have everything setup before the class begins.</p><p><strong>Please ensure the below hardware and software are available and set up:</strong></p><ul style="list-style:none;padding-left:0;"><li style="margin:8px 0;">&#10004;&#65039; Please ensure you are <strong>camera-ready</strong>, as this is a virtual session. You will need to appear on camera for certain segments, such as Debrief Sessions, VILT Sessions, and Q&amp;A</li><li style="margin:8px 0;">&#10060; Please <strong>do not</strong> open Adobe Connect via a browser</li><li style="margin:8px 0;">&#10004;&#65039; Corporate USB Headset (Not your own headset)</li><li style="margin:8px 0;">&#10004;&#65039; Please be available on Slack during the training</li></ul><div style="background-color:#f0f0f0;padding:12px;border-radius:5px;margin:15px 0;"><p style="margin:0;"><strong>Issues with Adobe Connect Application?</strong></p><p style="margin:5px 0;">- <a href="https://share.amazon.com/sites/kingfisher/_layouts/15/WopiFrame2.aspx?sourcedoc=%7b96D314DC-517F-4354-89D8-2A827B036478%7d&amp;file=Adobe_Connect_Troubleshooting_FAQ_Handout.pdf&amp;action=default">Troubleshoot Adobe Connect</a> | <a href="https://helpx.adobe.com/in/support/connect.html">24*7 Chat Support</a></p><p style="margin:5px 0;">- If you cannot install it due to laptop permissions, <a href="https://my.it.a2z.com/articles/access/administrator/admin-access-windows">request Admin Rights</a></p><p style="margin:5px 0;">- If you still struggle with installing Adobe Connect, <strong><a href="https://my.it.a2z.com">contact IT immediately</a></strong></p></div><hr style="border:none;border-top:1px solid #ddd;margin:20px 0;"><p>Please reach out, should there be any questions/concerns. I'm looking forward to having you in the training!</p><p style="margin-top:30px;"><strong>{trainer_name}</strong></p></div></body></html>"""

def generate_plain_text_body(skill_details, trainer_name, shift_details, adobe_link):
    return f"""Hello Everyone!

You have been invited to join a training session on {skill_details}

Your Trainer will be: {trainer_name}. Shift: {shift_details}

Please respond to this e-invite to confirm your attendance.

---

Link to Adobe Connect Room (to be opened via the Adobe Connect App only)

{adobe_link}

NOTE: Since all WW SPS-WE trainings may involve customer accounts, Slack is NOT an approved tool for sharing customer data. Please make sure Adobe Connect is installed and ready before your training begins.

Supervisors and Managers listed as optional attendees are added for visibility only. If this training isn't relevant to you, please skip it and/or decline the invite so it doesn't block your calendar.

---

First time using Adobe Connect?

Please ensure that the Adobe Connect application is installed on your laptop. If it isn't already installed, you can do so by following these steps:
Start Menu > Search for "Software Center" > Open Software Center > Search for "Adobe Connect" > Click "Install."

Once installed, please copy the room link, paste it in the application and click on Continue. Now select Guest as the option and type your full name. Post that, click on Enter Room.

Please join 5 - 10 minutes early to have everything setup before the class begins.

Finally, ensure that the below hardware and software are available and set up:

  [YES] Please ensure you are camera-ready, as this is a virtual session.
  [NO]  Please do NOT open Adobe Connect via a browser
  [YES] Corporate USB Headset (Not your own headset)
  [YES] Please be available on Slack during the training

---

Issues with Adobe Connect Application?
  - Troubleshoot Adobe Connect | 24*7 Chat Support
  - If you cannot install it due to laptop permissions, request Admin Rights
  - If you still struggle with installing Adobe Connect, contact IT immediately

---

Please reach out, should there be any questions/concerns. I'm looking forward to having you in the training!

{trainer_name}
"""

def generate_ics_content(subject, description, start_dt, end_dt, required_emails, optional_emails, organizer_name, organizer_email, location):
    uid = str(uuid.uuid4())
    dtstart = start_dt.strftime("%Y%m%dT%H%M%S")
    dtend = end_dt.strftime("%Y%m%dT%H%M%S")
    dtstamp = datetime.datetime.now().strftime("%Y%m%dT%H%M%SZ")
    desc_ics = description.replace("\\", "\\\\").replace("\n", "\\n").replace(",", "\\,").replace(";", "\\;")
    attendee_lines = ""
    for email in required_emails:
        email = email.strip()
        if email:
            attendee_lines += f"ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN={email}:mailto:{email}\n"
    for email in optional_emails:
        email = email.strip()
        if email:
            attendee_lines += f"ATTENDEE;ROLE=OPT-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN={email}:mailto:{email}\n"
    return f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//SPS WE-TD Training Invite Generator//EN
CALSCALE:GREGORIAN
METHOD:REQUEST
BEGIN:VEVENT
UID:{uid}
DTSTAMP:{dtstamp}
DTSTART:{dtstart}
DTEND:{dtend}
ORGANIZER;CN={organizer_name}:mailto:{organizer_email}
SUMMARY:{subject}
DESCRIPTION:{desc_ics}
LOCATION:{location}
{attendee_lines}STATUS:CONFIRMED
SEQUENCE:0
TRANSP:OPAQUE
X-MICROSOFT-CDO-BUSYSTATUS:BUSY
X-MICROSOFT-CDO-INTENDEDSTATUS:BUSY
X-MICROSOFT-CDO-ALLDAYEVENT:FALSE
X-MICROSOFT-CDO-IMPORTANCE:1
BEGIN:VALARM
TRIGGER:-PT15M
ACTION:DISPLAY
DESCRIPTION:Training Reminder
END:VALARM
END:VEVENT
END:VCALENDAR"""

def send_outlook_calendar_invite(subject, plain_body, required_emails, optional_emails, start_dt, end_dt):
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        appointment = outlook.CreateItem(1)
        appointment.MeetingStatus = 1
        appointment.Subject = subject
        appointment.Body = plain_body
        appointment.Start = start_dt.strftime("%Y-%m-%d %H:%M")
        appointment.End = end_dt.strftime("%Y-%m-%d %H:%M")
        appointment.ReminderSet = True
        appointment.ReminderMinutesBeforeStart = 15
        appointment.BusyStatus = 2
        if required_emails:
            for email in required_emails:
                email = email.strip()
                if email:
                    recipient = appointment.Recipients.Add(email)
                    recipient.Type = 1
        if optional_emails:
            for email in optional_emails:
                email = email.strip()
                if email:
                    recipient = appointment.Recipients.Add(email)
                    recipient.Type = 2
        appointment.Recipients.ResolveAll()
        appointment.Send()
        return True, "Calendar invite sent successfully!"
    except Exception as e:
        return False, f"Error sending invite: {str(e)}"
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

st.subheader("Preview and Send")

DEFAULT_DL = "griffin-btr@share.corp.amazon.com"

required_fields = {
    "Skill Name": skill_name,
    "Batch ID": batch_id,
    "Skill Details": skill_details,
    "Trainer Name": trainer_name,
    "Shift Details": shift_details,
    "Adobe Connect Link": adobe_connect_link,
    "Required Attendees": required_attendees,
    "Sender Email": sender_email
}

missing_fields = [k for k, v in required_fields.items() if not v.strip()]

if missing_fields:
    st.warning(f"Please fill in the following required fields: **{', '.join(missing_fields)}**")

if all(v.strip() for v in required_fields.values()):
    subject = generate_subject(skill_name, batch_id, start_date, end_date)
    html_body = generate_html_body(skill_details, trainer_name, shift_details, adobe_connect_link)
    plain_body = generate_plain_text_body(skill_details, trainer_name, shift_details, adobe_connect_link)
    duration_display = f"{format_date_display(start_date)} to {format_date_display(end_date)}"
    req_emails = process_attendees(required_attendees)
    opt_emails = process_attendees(optional_attendees)
    if DEFAULT_DL not in opt_emails:
        opt_emails.append(DEFAULT_DL)
    start_dt = datetime.datetime.combine(start_date, datetime.time(9, 0))
    end_dt = datetime.datetime.combine(start_date, datetime.time(18, 0))

    with st.expander("Preview Email Subject and Details", expanded=True):
        st.markdown(f"**Subject:** `{subject}`")
        st.markdown(f"**Duration:** {duration_display}")
        st.markdown(f"**Required Attendees:** {len(req_emails)} recipient(s)")
        for email in req_emails:
            st.markdown(f"  - `{email}`")
        st.markdown(f"**Optional Attendees:** {len(opt_emails)} recipient(s)")
        for email in opt_emails:
            st.markdown(f"  - `{email}`")
        st.markdown(f"**Signature:** {trainer_name}")

    with st.expander("Preview Email Body (HTML)", expanded=False):
        st.components.v1.html(html_body, height=600, scrolling=True)

    st.divider()
    st.subheader("Send Calendar Invite")

    if OUTLOOK_AVAILABLE:
        if st.button("Send Calendar Invite via Outlook", type="primary", use_container_width=False):
            with st.spinner("Sending calendar invite via Outlook..."):
                success, message = send_outlook_calendar_invite(subject, plain_body, req_emails, opt_emails, start_dt, end_dt)
            if success:
                st.success(message)
                st.balloons()
                st.markdown(f"""<div class="success-box"><strong>Invite Sent Successfully!</strong><br><br>Subject: <strong>{subject}</strong><br>Required: <strong>{len(req_emails)} recipient(s)</strong><br>Optional: <strong>{len(opt_emails)} recipient(s)</strong><br>Date: <strong>{start_date.strftime('%B %d, %Y')}</strong><br>Signature: <strong>{trainer_name}</strong></div>""", unsafe_allow_html=True)
            else:
                st.error(message)
    else:
        st.markdown("""<div class="step-box"><strong>How it works:</strong><br>1. Click the button below to download the calendar invite file<br>2. Double-click the downloaded <code>.ics</code> file<br>3. Outlook opens with a <strong>Meeting Invite</strong> with all attendees and details pre-filled<br>4. Click <strong>Send</strong> - Done!</div>""", unsafe_allow_html=True)
        st.markdown("")
        ics_content = generate_ics_content(subject=subject, description=plain_body, start_dt=start_dt, end_dt=end_dt, required_emails=req_emails, optional_emails=opt_emails, organizer_name=trainer_name, organizer_email=sender_email, location=adobe_connect_link)
        safe_skill = re.sub(r'[^a-zA-Z0-9]', '_', skill_name)[:30]
        ics_filename = f"Training_Invite_{safe_skill}.ics"
        col_dl, col_space = st.columns([1, 2])
        with col_dl:
            st.download_button(label="Download Calendar Invite (.ics)", data=ics_content, file_name=ics_filename, mime="text/calendar", type="primary", use_container_width=True)

st.divider()
mode_text = "Local Mode (Outlook COM)" if OUTLOOK_AVAILABLE else "Cloud Mode (Calendar ICS)"
st.markdown(f"""<div style="text-align:center;color:#888;font-size:0.85rem;">SPS WE-TD Training Invitation Template Generator v3.0 | {mode_text}</div>""", unsafe_allow_html=True)
