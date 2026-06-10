# Training Invitation Generator (SPS-WE IND)

## Overview
A Streamlit web app that generates training invitation calendar invites.

## Two Modes:
1. **Local Mode (Windows + Outlook):** Sends calendar invites directly via Outlook COM
2. **Cloud Mode (Any platform):** Downloads .ics file that user opens in their Outlook

## Local Setup (Windows)
```bash
pip install streamlit pywin32
streamlit run training_invite_app.py
```

## Cloud Deployment (Streamlit Community Cloud)
1. Push this repo to GitHub
2. Go to https://share.streamlit.io
3. Connect your GitHub repo
4. Deploy - get a shareable URL!

## Files
- `training_invite_app.py` - Main app
- `requirements.txt` - Python dependencies (cloud-compatible)
- `.streamlit/config.toml` - Theme and server config
- `README.md` - This file
