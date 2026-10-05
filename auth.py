"""Minimal login page with a "stay signed in" cookie.

Users and passwords live in Streamlit Secrets (never in the repo):

    [users]
    manager = "a-password"
    soumava = "another-password"

Fails closed: with no [users] configured nobody gets in, so a forgotten secret cannot leave the app open.

Stay signed in: after a correct login the browser stores a signed token (cookie) for REMEMBER_DAYS days.
On the next visit a valid token skips the login page. The token is signed with that user's password, so
changing a password in Secrets signs that user out everywhere.
"""
import hashlib
import hmac
import time
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

REMEMBER_DAYS = 30
COOKIE = "lead_auth"
# Reads the cookie inside the browser and hands it to Python (see components/cookie_reader/index.html)
_cookie_reader = components.declare_component("cookie_reader", path=str(Path(__file__).parent / "components" / "cookie_reader"))


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


# ------------------------------------------------------------ remember-me token

def _sign(key_name, password, expiry):
    key = hashlib.sha256(f"lead-automator|{key_name}|{password}".encode()).digest()
    return hmac.new(key, f"{key_name}~{expiry}".encode(), hashlib.sha256).hexdigest()


def make_token(users, username, now=None):
    name = username.strip().casefold()
    expiry = int((now if now is not None else time.time()) + REMEMBER_DAYS * 86400)
    return f"{name}~{expiry}~{_sign(name, users[name][1], expiry)}"


def verify_token(users, token, now=None):
    """Display name for a valid, unexpired token, else None."""
    try:
        name, expiry, sig = (token or "").rsplit("~", 2)
        entry = users.get(name)
        if entry and int(expiry) > (now if now is not None else time.time()) \
                and hmac.compare_digest(sig, _sign(name, entry[1], int(expiry))):
            return entry[0]
    except ValueError:
        pass
    return None


def _cookie_from_browser():
    try:
        return st.context.cookies.get(COOKIE)
    except Exception:  # older Streamlit / no request context
        return None


def _queue_cookie(value, max_age):
    st.session_state["_cookie_js"] = (value, max_age)


def apply_cookie_changes():
    """Write any pending cookie change into the browser. Call once at the end of every page render."""
    pending = st.session_state.pop("_cookie_js", None)
    if not pending:
        return
    value, max_age = pending
    components.html(f"""<script>
      var c = "{COOKIE}={value}; max-age={max_age}; path=/; SameSite=Lax" + (location.protocol === "https:" ? "; Secure" : "");
      try {{ window.parent.document.cookie = c; }} catch (e) {{ document.cookie = c; }}
    </script>""", height=0)


# ------------------------------------------------------------ page flow

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

    # st.context.cookies is read once when the page opens, so after signing out it still shows the
    # old cookie. The flag stops that stale cookie from signing the user straight back in.
    if not st.session_state.get("signed_out"):
        name = verify_token(users, _cookie_from_browser())  # works when the server receives cookies (local runs)
        if name:
            st.session_state["user"] = name
            return name

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
            st.session_state["signed_out"] = False
            _queue_cookie(make_token(users, name), REMEMBER_DAYS * 86400)
            st.rerun()
        time.sleep(1)  # slows down password guessing
        st.error("Wrong username or password.")
    if not st.session_state.get("signed_out"):
        # Read the cookie in the browser instead: it reaches us even when the hosting proxy hides cookies from the
        # server. None = the browser has not answered yet, "" = no cookie. A valid token signs the user in.
        name = verify_token(users, _cookie_reader(name=COOKIE, key="cookie_reader", default=None))
        if name:
            st.session_state["user"] = name
            st.rerun()
    apply_cookie_changes()  # clears the cookie after a sign out
    st.stop()


def sign_out():
    st.session_state.pop("user", None)
    st.session_state["signed_out"] = True
    _queue_cookie("", 0)  # max-age=0 deletes the cookie
