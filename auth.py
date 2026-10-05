"""Minimal login page. Users and passwords live in Streamlit Secrets (never in the repo):

    [users]
    manager = "a-password"
    soumava = "another-password"

Fails closed: with no [users] configured nobody gets in, so a forgotten secret cannot leave the app open.
"""
import hmac
import time

import streamlit as st


def _users():
    """{lowercase username: (display name, password)} from secrets."""
    try:
        raw = st.secrets.get("users")
        return {str(k).strip().casefold(): (str(k).strip(), str(v)) for k, v in dict(raw).items()} if raw else {}
    except Exception:  # no secrets file at all
        return {}


def check(users, username, password):
    """Display name when the credentials are valid, else None."""
    entry = users.get((username or "").strip().casefold())
    if entry and hmac.compare_digest(entry[1].encode(), (password or "").encode()):
        return entry[0]
    return None


def require_login():
    """Show the login page and stop the script until a valid user signs in. Returns the user's name."""
    users = _users()
    if not users:
        st.title("Lead Automator")
        st.error("No users are set up yet. Add a [users] section to the app's Secrets.")
        st.stop()

    current = st.session_state.get("user")
    if current and current.casefold() in users:
        return current

    st.title("Lead Automator")
    st.caption("Sign in to continue.")
    with st.form("login_form", border=False):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Sign in", type="primary")
    if submitted:
        name = check(users, username, password)
        if name:
            st.session_state["user"] = name
            st.rerun()
        time.sleep(1)  # slows down password guessing
        st.error("Wrong username or password.")
    st.stop()


def sign_out():
    st.session_state.pop("user", None)
