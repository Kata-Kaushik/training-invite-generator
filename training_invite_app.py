
import streamlit as st
import datetime
import re
import platform
import uuid
import urllib.parse
from io import BytesIO

# Try to import Windows-specific modules
OUTLOOK_AVAILABLE = False
try:
    if platform.system() == "Windows":
        import win32com.client
        import pythoncom
        OUTLOOK_AVAILABLE = True
except ImportError:
    pass

# ============================================================
# TRAINING INVITATION GENERATOR & SENDER
# SPS-WE IND - Streamlit + Outlook COM (local) + Mailto (cloud)
# ============================================================

st.set_page_config(
    page_title="Training Invitation Generator",
    page_icon="mail",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: bold;
        color: #1a73e8;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #555;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
    }
    .warning-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #fff3cd;
        border: 1px solid #ffc107;
        color: #856404;
    }
    .outlook-btn {
        display: inline-block;
        padding: 0.75rem 2rem;
        background-color: #0078d4;
        color: white !important;
        text-decoration: none;
        border-radius: 5px;
        font-size: 1.1rem;
        font-weight: bold;
        text-align: center;
        margin: 10px 0;
    }
    .outlook-btn:hover {
        background-color: #005a9e;
        color: white !important;
        text-decoration: none;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# HEADER
# ============================================================
st.markdown('<div class="main-header">Training Invitation Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">SPS-WE IND | Auto-generate and send training calendar invites</div>', unsafe_allow_html=True)

# Show mode indicator
if OUTLOOK_AVAILABLE:
    st.success("**Mode: Local (Outlook Direct Send)** - Calendar invites will be sent directly from your Outlook")
else:
    st.info("**Mode: Cloud** - Click the generated button to open a pre-filled invite in your Outlook")

st.divider()

# ============================================================
# INPUT FORM
# ============================================================
col1, col2 = st.columns(2)

with col1:
    st.subheader("Training Details")

    skill_name = st.text_input(
        "Skill / Training Name *",
        placeholder="e.g., NH MKT Training",
        help="The name of the training skill/program"
    )

    batch_id = st.text_input(
        "Batch ID *",
        placeholder="e.g., SPSSELLER-LNH-IND-BLR-2024-04-22-X-MKT33",
        help="Unique batch identifier"
    )

    skill_details = st.text_area(
        "Skill Details (Topic Description) *",
        placeholder="e.g., New Hire Marketing Training - Module 3: Campaign Management",
        help="Detailed description of what the training covers",
        height=80
    )

    trainer_name = st.text_input(
        "Trainer Name *",
        placeholder="e.g., John Doe",
        help="Full name of the trainer"
    )

    shift_details = st.text_input(
        "Shift Details and WOs *",
        placeholder="e.g., Shift A (9:00 AM - 6:00 PM IST), WO: 12345",
        help="Shift timing and Work Order details"
    )

    adobe_connect_link = st.text_input(
        "Adobe Connect Room Link *",
        placeholder="e.g., https://amazon.adobeconnect.com/room-name/",
        help="Full URL to the Adobe Connect training room"
    )

with col2:
    st.subheader("Schedule and Recipients")

    col_date1, col_date2 = st.columns(2)
    with col_date1:
        start_date = st.date_input(
            "Start Date *",
            value=datetime.date.today() + datetime.timedelta(days=7),
            help="Training start date"
        )
    with col_date2:
        end_date = st.date_input(
            "End Date *",
            value=datetime.date.today() + datetime.timedelta(days=37),
            help="Training end date"
        )

    col_time1, col_time2 = st.columns(2)
    with col_time1:
        start_time = st.time_input(
            "Daily Start Time *",
            value=datetime.time(9, 0),
            help="Daily training start time"
        )
    with col_time2:
        end_time = st.time_input(
            "Daily End Time *",
            value=datetime.time(18, 0),
            help="Daily training end time"
        )

    required_attendees = st.text_area(
        "Required Attendees (Logins/Emails) *",
        placeholder="Enter email addresses separated by semicolons (;)\ne.g., user1@amazon.com; user2@amazon.com",
        help="Logins attending the training - separate multiple emails with semicolons",
        height=100
    )

    optional_attendees = st.text_area(
        "Optional Attendees (Supervisors/STMs)",
        placeholder="Enter email addresses separated by semicolons (;)\ne.g., manager1@amazon.com; stm1@amazon.com",
        help="Supervisors and managers for visibility only",
        height=80
    )

    sender_name = st.text_input(
        "Your Name (Sender) *",
        placeholder="e.g., John Doe",
        help="Your name for the signature"
    )

    sender_signature = st.text_area(
        "Your Signature (Optional)",
        placeholder="e.g., John Doe | SPS-WE Training Team | Slack: @johndoe",
        help="Custom signature block",
        height=60
    )

st.divider()

# ============================================================
# TEMPLATE GENERATION FUNCTIONS
# ============================================================

def format_date_display(date_obj):
    """Format date for display in the invitation"""
    day = date_obj.day
    if 4 <= day <= 20 or 24 <= day <= 30:
        suffix = "th"
    else:
        suffix = ["st", "nd", "rd"][day % 10 - 1]
    return date_obj.strftime(f"%B {day}{suffix}")


def generate_subject(skill, batch, start, end):
    """Generate the email subject line"""
    duration = f"{format_date_display(start)} to {format_date_display(end)}"
    return f"{skill} | {batch} | {duration}"


def generate_html_body(skill_details, trainer_name, shift_details, adobe_link, sender_name, sender_sig):
    """Generate the HTML email body for preview"""
    signature = sender_sig if sender_sig else sender_name

    html_body = f"""<html>
<body style="font-family: 'Segoe UI', Arial, sans-serif; font-size: 14px; color: #333; line-height: 1.6;">
<div style="max-width: 700px; margin: 0 auto; padding: 20px;">

<p style="font-size: 15px;">Hello Everyone!</p>

<p>You have been invited to join a training session on <strong>{skill_details}</strong></p>

<p>Your Trainer will be: <strong>{trainer_name}</strong>. Shift: <strong>{shift_details}</strong></p>

<p><strong>Please respond to this e-invite to confirm your attendance.</strong></p>

<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">

<p><strong>Link to Adobe Connect Room</strong> (to be opened via the Adobe Connect App only)</p>
<p><a href="{adobe_link}" style="color: #1a73e8; font-weight: bold;">{adobe_link}</a></p>

<div style="background-color: #fff3cd; border: 1px solid #ffc107; border-radius: 5px; padding: 12px; margin: 15px 0;">
<p style="margin: 0;"><strong>WARNING - Note:</strong> Since all WW SPS-WE trainings may involve customer accounts, Slack is not an approved tool for sharing customer data. Please make sure Adobe Connect is installed and ready before your training begins.</p>
<p style="margin: 8px 0 0 0; font-style: italic;">Supervisors and Managers listed as optional attendees are added for visibility only. If this training is not relevant to you, please skip it and/or decline the invite so it does not block your calendar.</p>
</div>

<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">

<p><strong>First time using Adobe Connect?</strong></p>
<p>Please ensure that the Adobe Connect application is installed on your laptop. If it is not already installed, you can do so by following these steps:</p>
<p style="background-color: #e8f4fd; padding: 10px; border-radius: 5px; border-left: 4px solid #1a73e8;">
<strong>Start Menu &gt; Search for "Software Center" &gt; Open Software Center &gt; Search for "Adobe Connect" &gt; Click "Install"</strong>
</p>

<p>Once installed, please copy the room link, paste it in the application and click on <strong>Continue</strong>. Now select <strong>Guest</strong> as the option and type your full name. Post that, click on <strong>Enter Room</strong>. Please join <strong>5-10 minutes early</strong> to have everything setup before the class begins.</p>

<p><strong>Please ensure the below hardware and software are available and set up:</strong></p>

<ul style="list-style: none; padding-left: 0;">
<li style="margin: 8px 0;">&#10004;&#65039; Please ensure you are <strong>camera-ready</strong>, as this is a virtual session. You will need to appear on camera for certain segments, such as Debrief Sessions, VILT Sessions, and Q&amp;A</li>
<li style="margin: 8px 0;">&#10060; Please <strong>do not</strong> open Adobe Connect via a browser</li>
<li style="margin: 8px 0;">&#10004;&#65039; Corporate USB Headset (Not your own headset)</li>
<li style="margin: 8px 0;">&#10004;&#65039; Please be available on Slack during the training</li>
</ul>

<div style="background-color: #f0f0f0; padding: 12px; border-radius: 5px; margin: 15px 0;">
<p style="margin: 0;"><strong>Issues with Adobe Connect Application?</strong></p>
<p style="margin: 5px 0;">- <a href="https://share.amazon.com/sites/kingfisher/_layouts/15/WopiFrame2.aspx?sourcedoc=%7b96D314DC-517F-4354-89D8-2A827B036478%7d&file=Adobe_Connect_Troubleshooting_FAQ_Handout.pdf&action=default">Troubleshoot Adobe Connect</a> | <a href="https://helpx.adobe.com/in/support/connect.html">24*7 Chat Support</a></p>
<p style="margin: 5px 0;">- If you cannot install it due to laptop permissions, <a href="https://my.it.a2z.com/articles/access/administrator/admin-access-windows">request Admin Rights</a></p>
<p style="margin: 5px 0;">- If you still struggle with installing Adobe Connect, <strong><a href="https://my.it.a2z.com">contact IT immediately</a></strong></p>
</div>

<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">

<p>Please reach out, should there be any questions/concerns. I'm looking forward to having you in the training!</p>

<p style="margin-top: 30px;"><strong>{signature}</strong></p>

</div>
</body>
</html>"""
    return html_body


def generate_plain_text_body(skill_details, trainer_name, shift_details, adobe_link, sender_name, sender_sig):
    """Generate plain text email body for mailto and Outlook COM"""
    signature = sender_sig if sender_sig else sender_name

    body = f"""Hello Everyone!

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

{signature}
"""
    return body


def generate_mailto_link(subject, body, required_emails, optional_emails):
    """Generate a mailto: link that opens Outlook with pre-filled content"""
    to_field = ";".join([e.strip() for e in required_emails if e.strip()])
    cc_field = ";".join([e.strip() for e in optional_emails if e.strip()])

    # URL encode the components
    params = {
        "subject": subject,
        "body": body,
    }
    if cc_field:
        params["cc"] = cc_field

    query_string = urllib.parse.urlencode(params, quote_via=urllib.parse.quote)
    mailto_link = f"mailto:{to_field}?{query_string}"

    return mailto_link


def send_outlook_calendar_invite(subject, plain_body, required_emails, optional_emails, start_dt, end_dt):
    """Send calendar invite via Outlook COM automation (Windows only)"""
    try:
        pythoncom.CoInitialize()

        outlook = win32com.client.Dispatch("Outlook.Application")
        appointment = outlook.CreateItem(1)  # 1 = olAppointmentItem

        appointment.MeetingStatus = 1  # olMeeting
        appointment.Subject = subject
        appointment.Body = plain_body
        appointment.Start = start_dt.strftime("%Y-%m-%d %H:%M")
        appointment.End = end_dt.strftime("%Y-%m-%d %H:%M")
        appointment.ReminderSet = True
        appointment.ReminderMinutesBeforeStart = 15
        appointment.BusyStatus = 2  # olBusy

        if required_emails:
            for email in required_emails:
                email = email.strip()
                if email:
                    recipient = appointment.Recipients.Add(email)
                    recipient.Type = 1  # olRequired

        if optional_emails:
            for email in optional_emails:
                email = email.strip()
                if email:
                    recipient = appointment.Recipients.Add(email)
                    recipient.Type = 2  # olOptional

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


# ============================================================
# PREVIEW & SEND
# ============================================================

st.subheader("Preview and Send")

# Validate required fields
required_fields = {
    "Skill Name": skill_name,
    "Batch ID": batch_id,
    "Skill Details": skill_details,
    "Trainer Name": trainer_name,
    "Shift Details": shift_details,
    "Adobe Connect Link": adobe_connect_link,
    "Required Attendees": required_attendees,
    "Sender Name": sender_name
}

missing_fields = [k for k, v in required_fields.items() if not v.strip()]

if missing_fields:
    st.warning(f"Please fill in the following required fields: **{', '.join(missing_fields)}**")

# Generate preview and action buttons
if all(v.strip() for v in required_fields.values()):

    # Generate subject
    subject = generate_subject(skill_name, batch_id, start_date, end_date)

    # Generate bodies
    html_body = generate_html_body(
        skill_details, trainer_name, shift_details,
        adobe_connect_link, sender_name, sender_signature
    )
    plain_body = generate_plain_text_body(
        skill_details, trainer_name, shift_details,
        adobe_connect_link, sender_name, sender_signature
    )

    # Duration display
    duration_display = f"{format_date_display(start_date)} to {format_date_display(end_date)}"

    # Parse email lists
    req_emails = [e.strip() for e in required_attendees.split(";") if e.strip()]
    opt_emails = [e.strip() for e in optional_attendees.split(";") if e.strip()] if optional_attendees else []

    # Create datetime objects
    start_dt = datetime.datetime.combine(start_date, start_time)
    end_dt = datetime.datetime.combine(start_date, end_time)

    # Preview section
    with st.expander("Preview Email Subject and Details", expanded=True):
        st.markdown(f"**Subject:** `{subject}`")
        st.markdown(f"**Duration:** {duration_display}")
        st.markdown(f"**Time:** {start_time.strftime('%I:%M %p')} - {end_time.strftime('%I:%M %p')}")
        st.markdown(f"**Required Attendees:** {len(req_emails)} recipient(s) - {', '.join(req_emails)}")
        if opt_emails:
            st.markdown(f"**Optional Attendees:** {len(opt_emails)} recipient(s) - {', '.join(opt_emails)}")

    with st.expander("Preview Email Body (HTML)", expanded=False):
        st.components.v1.html(html_body, height=600, scrolling=True)

    st.divider()

    # ============================================================
    # ACTION BUTTONS
    # ============================================================

    st.subheader("Send Invite")

    if OUTLOOK_AVAILABLE:
        # LOCAL MODE: Direct Outlook Send
        col_btn1, col_btn2, col_spacer = st.columns([1, 1, 2])

        with col_btn1:
            if st.button("Send via Outlook (Direct)", type="primary", use_container_width=True):
                with st.spinner("Sending calendar invite via Outlook..."):
                    success, message = send_outlook_calendar_invite(
                        subject, plain_body, req_emails, opt_emails, start_dt, end_dt
                    )
                if success:
                    st.success(message)
                    st.balloons()
                else:
                    st.error(message)

        with col_btn2:
            # Also offer the mailto option locally
            mailto_link = generate_mailto_link(subject, plain_body, req_emails, opt_emails)
            st.markdown(
                f'<a href="{mailto_link}" class="outlook-btn" target="_blank">Open in Outlook (Email)</a>',
                unsafe_allow_html=True
            )

    else:
        # CLOUD MODE: Mailto link to open user's Outlook
        st.markdown("""
        <div class="info-box">
        <strong>How it works:</strong> Click the button below to open your Outlook with the invitation 
        pre-filled (To, CC, Subject, and Body). Just review and click Send!
        </div>
        """, unsafe_allow_html=True)

        st.markdown("")

        # Generate mailto link
        mailto_link = generate_mailto_link(subject, plain_body, req_emails, opt_emails)

        # Display as a prominent clickable button
        st.markdown(
            f"""
            <div style="text-align: center; margin: 20px 0;">
                <a href="{mailto_link}" class="outlook-btn" target="_blank">
                    &#9993; Open in Outlook - Send Invite
                </a>
            </div>
            <p style="text-align: center; color: #666; font-size: 0.9rem;">
                Clicking this button will open your default email client (Outlook) with all fields pre-filled.<br>
                Just review the content and click <strong>Send</strong>!
            </p>
            """,
            unsafe_allow_html=True
        )

        st.markdown("")

        # Also provide a copy-paste fallback
        with st.expander("Alternative: Copy content manually", expanded=False):
            st.markdown("**Subject (copy this):**")
            st.code(subject, language=None)
            st.markdown("**To (copy this):**")
            st.code(";".join(req_emails), language=None)
            if opt_emails:
                st.markdown("**CC (copy this):**")
                st.code(";".join(opt_emails), language=None)
            st.markdown("**Body (copy this):**")
            st.code(plain_body, language=None)

# ============================================================
# FOOTER
# ============================================================
st.divider()
mode_text = "Local Mode (Outlook COM)" if OUTLOOK_AVAILABLE else "Cloud Mode (Mailto)"
st.markdown(f"""
<div style="text-align: center; color: #888; font-size: 0.85rem;">
    Training Invitation Generator v2.1 | SPS-WE IND | {mode_text}
</div>
""", unsafe_allow_html=True)
