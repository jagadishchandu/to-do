"""Standalone helper: a persistent, navigable embedded browser window (a
content window plus a small address-bar toolbar window) for viewing and
logging into any portal, Chrome-tab style.

Runs as a subprocess because the browser engine needs the main thread, which
the main app's Tkinter loop already owns — so this isn't literally a tab
inside the main window, it's a separate window the app launches/manages.

Uses a fixed, shared profile directory so cookies with an explicit expiry
(most "remember me" / long-lived logins) persist across separate launches —
you often won't need to sign in again next time. Session-only cookies (many
SSO logins, by design) still won't survive a restart; that's normal browser
behavior, not a limitation of this tool.

Usage: python portal_browser_helper.py <initial_url> <profile_dir>
"""
import json
import sys
import threading

import webview


class _Api:
    """NOTE: deliberately holds no reference to any pywebview Window object.
    Storing a Window as an attribute on a js_api instance was confirmed (by
    testing) to make pywebview's JS-bridge introspection recurse into the
    window's native COM/accessibility properties and blow the recursion
    limit. The navigate/back/forward/reload methods are set as plain bound
    functions from main() instead — closures over the window, not attributes
    of this object."""
    pass


TOOLBAR_HTML = """
<html><body style="margin:0;padding:6px;background:#181824;">
<div style="display:flex;gap:6px;align-items:center;font-family:Segoe UI,sans-serif;">
  <button onclick="pywebview.api.back()" style="{btn}">&#8592;</button>
  <button onclick="pywebview.api.forward()" style="{btn}">&#8594;</button>
  <button onclick="pywebview.api.reload()" style="{btn}">&#8635;</button>
  <input id="addr" type="text" style="flex:1;padding:8px 10px;border-radius:8px;
      border:1px solid #33334a;background:#22222f;color:#f2f2f7;font-size:13px;"
      onkeydown="if(event.key==='Enter'){{pywebview.api.navigate(document.getElementById('addr').value);}}" />
  <button onclick="pywebview.api.navigate(document.getElementById('addr').value)"
      style="{btn}background:#7c6cff;color:white;">Go</button>
</div>
</body></html>
""".format(btn="padding:8px 12px;border-radius:8px;border:1px solid #33334a;background:#22222f;color:#f2f2f7;cursor:pointer;")


def main():
    if len(sys.argv) < 3:
        print(json.dumps({"error": "usage: portal_browser_helper.py <url> <profile_dir>"}))
        return
    initial_url = sys.argv[1].strip()
    profile_dir = sys.argv[2].strip()
    if "://" not in initial_url:
        initial_url = "https://" + initial_url

    api = _Api()
    content_window = webview.create_window("Portal Browser", initial_url, width=1100, height=780)
    toolbar_window = webview.create_window(
        "Address bar", html=TOOLBAR_HTML, js_api=api, width=1100, height=120,
    )

    def navigate(url: str):
        url = (url or "").strip()
        if not url:
            return False
        if "://" not in url:
            url = "https://" + url
        content_window.load_url(url)
        return True

    def back():
        content_window.evaluate_js("history.back()")
        return True

    def forward():
        content_window.evaluate_js("history.forward()")
        return True

    def reload_page():
        content_window.evaluate_js("location.reload()")
        return True

    api.navigate = navigate
    api.back = back
    api.forward = forward
    api.reload = reload_page

    def sync_address_bar():
        try:
            url = content_window.get_current_url() or ""
            toolbar_window.evaluate_js(f"document.getElementById('addr').value = {json.dumps(url)}")
        except Exception:
            pass

    content_window.events.loaded += sync_address_bar

    webview.start(gui="edgechromium", debug=False, private_mode=False, storage_path=profile_dir)


if __name__ == "__main__":
    main()
