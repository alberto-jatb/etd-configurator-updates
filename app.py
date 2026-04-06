"""
ETD Configurator - Main application window.
"""

import re
import os
import queue
import subprocess
import sys
from tkinter import messagebox, filedialog

import customtkinter as ctk

from geo_data import GEO_REGIONS, GEO_IPS, GEO_OUTBOUND_HOSTS
from ps_generator import generate_script, generate_steps
from executor import run_script
from version import VERSION, VERSION_DATE


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
IP_RE    = re.compile(r"^\d{1,3}(\.\d{1,3}){3}(/\d{1,2})?$")


def _valid_email(s):
    return bool(EMAIL_RE.match(s.strip()))


def _parse_ips(text):
    """Parse IPs from a multiline or comma-separated text block."""
    raw = re.split(r"[\n,;]+", text)
    return [ip.strip() for ip in raw if ip.strip()]


# ---------------------------------------------------------------------------
# Help texts
# ---------------------------------------------------------------------------

HELP_TEXTS = {
    "admin_upn": (
        "The User Principal Name (UPN) of the Exchange Online admin account "
        "used to connect and run configuration commands.\n\n"
        "Example: admin@contoso.onmicrosoft.com\n\n"
        "This account must have Exchange Administrator permissions."
    ),
    "deployment_mode": (
        "Select the ETD deployment mode for your organization:\n\n"
        "Journaling: Exchange sends copies of all emails to ETD via a journal "
        "rule. ETD analyzes them passively.\n\n"
        "ETD Inline: Exchange routes email traffic through ETD for active "
        "filtering. ETD is in the mail flow path. Requires configuring "
        "inbound and/or outbound connectors and transport rules."
    ),
    "geo_region": (
        "The geographic region where your ETD instance is deployed.\n\n"
        "This determines the IP addresses and hostnames used for connectors "
        "and transport rules.\n\n"
        "Available regions: North America, Europe, India, Australia, "
        "United Arab Emirates, Beta, Government."
    ),
    "journal_address": (
        "The email address of the ETD journaling mailbox.\n\n"
        "Exchange sends a copy of all emails to this address for ETD "
        "to analyze. This address is provided by Cisco ETD during onboarding.\n\n"
        "Example: etd-journal-abc123@us.etd.cisco.com"
    ),
    "notification_alert": (
        "An email address where journal delivery failure notifications "
        "will be sent.\n\n"
        "If a journaled message cannot be delivered, Exchange Online sends "
        "a Non-Delivery Report (NDR) to this address."
        "If the address is already set, leave this field blank"
    ),
    "seg_in_front": (
        "Enable this if a Cisco Secure Email Gateway (SEG) is deployed "
        "in front of Exchange Online.\n\n"
        "This is need for ETD to discover the real source IP."
    ),
    "seg_header": (
        "The name of the custom header that the SEG inserts into emails "
        "to identify their origin.\n\n"
        "This header is used by transport rules to identify messages "
        "already processed by the SEG.\n\n"
        "Example: X-IronPort-RemoteIP"
    ),
    "active_flows": (
        "Select which email flows to configure for ETD Inline mode:\n\n"
        "Inbound: Emails from the internet through ETD. Creates an inbound "
        "connector and transport rules (bypass spam, quarantine, junk).\n\n"
        "Outbound: Emails from your users to the internet through ETD. "
        "Creates an outbound connector and a tag transport rule.\n\n"
        "Internal: Emails between internal users through ETD. Creates a "
        "journal rule scoped to internal traffic."
    ),
    "inbound_ips": (
        "The ETD IP addresses used for the inbound connector and bypass rule.\n\n"
        "Auto-populated from the selected GEO region. These IPs represent "
        "the ETD scanning cluster that relays inbound email to Exchange.\n\n"
        "The inbound connector trusts email from these IPs, and the bypass "
        "rule skips EOP spam filtering for messages from them."
    ),
    "smarthost": (
        "The FQDN of the ETD outbound smart host.\n\n"
        "Exchange uses this address to route outbound email to ETD for "
        "scanning before delivery to the internet.\n\n"
        "Auto-filled based on the selected GEO region. Can be overridden.\n"
        "Example: out.us.etd.cisco.com"
    ),
    "xpass": (
        "The authentication token for the X-CSE-ETD-OUTBOUND-AUTH header.\n\n"
        "This header is inserted by the outbound transport rule and used "
        "by ETD to verify the email came from your Exchange tenant.\n\n"
        "This value must match the token configured in your ETD portal.\n"
        "Default: 12345 — change this to your actual token."
    ),
    "outbound_rule": (
        "Controls whether the ETD Outbound Tag transport rule is created "
        "in Enabled or Disabled state.\n\n"
        "Disabled is useful for testing: the rule exists but does not "
        "route outbound email through ETD yet. You can enable it later "
        "from the Exchange admin center or by running Install again."
    ),
    "internal_journal": (
        "The ETD journaling address for the internal email flow.\n\n"
        "Exchange sends a copy of all internal emails (between users in "
        "your organization) to this address for ETD analysis.\n\n"
        "Typically the same journaling address used in Journaling mode."
    ),
    "step_by_step": (
        "Execute the installation one step at a time.\n\n"
        "After each step completes you are asked whether to continue or stop.\n\n"
        "Useful for:\n"
        "  - Testing each component individually\n"
        "  - Troubleshooting a specific step\n"
        "  - Verifying each Exchange object before proceeding"
    ),
}

# ---------------------------------------------------------------------------
# User manual
# ---------------------------------------------------------------------------

USER_MANUAL = """\
CISCO SECURE EMAIL THREAT DEFENSE
Exchange Online Configurator  —  User Manual
Version {version}  ·  {date}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OVERVIEW
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
This tool automates the configuration of Microsoft Exchange Online to
work with Cisco Secure Email Threat Defense (ETD).

It connects to your Exchange Online tenant using PowerShell and creates
the necessary connectors, transport rules, and journal rules based on
your deployment mode and preferences.

Requirements:
  • PowerShell 7 (pwsh) installed on the local machine
  • Exchange Online PowerShell module (EXO V3)
      Install-Module -Name ExchangeOnlineManagement
  • An Exchange Online admin account with Exchange Administrator role

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
O365 CONNECTION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Admin UPN
  Enter the User Principal Name of the Exchange Online administrator.
  This account is used to connect to Exchange Online via PowerShell.

  Example:  admin@contoso.onmicrosoft.com

  Required role:  Exchange Administrator

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DEPLOYMENT MODE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Journaling Mode
  Exchange sends a BCC copy of all emails to the ETD journaling address.
  ETD analyzes emails passively without being in the mail flow.

  What gets configured:
    • Journaling NDR address  (admin UPN as bounce notification target)
    • Outbound connector      (routes journal emails directly via MX)
    • Journal rule            (sends all email copies to ETD)

  Use this mode when ETD is deployed in monitoring / BCC mode.

ETD Inline Mode
  Exchange routes email through ETD for active inline filtering.
  ETD is in the mail flow path and can block or quarantine messages.

  Configure which flows to enable:
    Inbound, Outbound, Internal — or any combination.

  Use this mode when ETD is deployed as an inline MTA gateway.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
JOURNALING MODE — OPTIONS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
GEO Region
  The geographic region of your ETD deployment. Determines which ETD
  infrastructure is referenced in connector and rule configuration.

Journaling Address
  The ETD email address that receives journaled email copies.
  Provided by Cisco during ETD onboarding.
  Example:  etd-journal-abc123@us.etd.cisco.com

Notification Alert Email
  An email address that receives Non-Delivery Reports if journaling
  fails. Typically the Exchange admin or a shared distribution list.

SEG in Front of O365
  Enable if a Cisco Secure Email Gateway (SEG) is deployed between
  the internet and Exchange Online. When enabled, provide the SEG
  header name used to identify messages already processed by the SEG.

SEG Header Name
  The name of the custom header inserted by the SEG.
  Example:  X-IronPort-RemoteIP

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ETD INLINE MODE — OPTIONS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
GEO Region
  Select your ETD geographic region. This auto-populates the ETD
  inbound IP addresses and the outbound smart host FQDN.

Active Flows

  Inbound
    Handles emails arriving from the internet through ETD.
    Creates:
      • Inbound Connector         trusts email from ETD IPs
      • Cisco Secure Email Threat Defense Bypass Spam Filter    bypasses EOP spam filter for ETD IPs
      • Cisco Secure Email Threat Defense Quarantine Rule       quarantines X-CSE-Quarantine messages
      • Cisco Secure Email Threat Defense Junk Rule             marks X-CSE-Junk messages (SCL 9)

  Outbound
    Handles emails sent by your users to the internet through ETD.
    Creates:
      • Outbound Connector        routes outbound mail to ETD smart host
      • Cisco Secure Email Threat Defense Outbound Tag Rule     adds X-CSE-ETD-OUTBOUND-AUTH header
                                  and routes via the outbound connector

    SmartHost
      The ETD outbound relay FQDN. Auto-filled from the GEO region.

    X-CSE-ETD-OUTBOUND-AUTH Value
      Authentication token for ETD. Must match the value in your ETD
      portal. Default is 12345 — change to your actual token.

    Outbound Tag Rule
      Choose whether to create the rule Enabled or Disabled.
      Disabled allows testing: the rule exists but does not yet route
      outbound email through ETD.

      WARNING: If the rule is set to Enabled at install time, it takes
      effect immediately and will affect live outbound mail flow.
      Recommendation: install with the rule Disabled, verify the
      connector is working correctly, then enable it manually.

  Internal
    Handles emails between users in your organization through ETD.
    Creates:
      • Journaling NDR address configuration
      • Journal Rule (Internal scope)  sends internal email copies to ETD

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ACTIONS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Verify
  Checks whether each ETD component is already configured in your
  Exchange tenant. Does not make any changes. Shows CONFIGURED or
  NOT CONFIGURED status for each object in the output console.

Install
  Creates all connectors, rules, and settings based on your
  configuration. Skips objects that already exist.
  On failure: offers to roll back all changes made during this run.

Remove
  Deletes all ETD-related connectors, transport rules, and journal
  rules from your Exchange tenant. Use this to cleanly uninstall.

Export .ps1
  Saves the generated PowerShell script to a .ps1 file without
  executing it. Useful for reviewing, running manually, or submitting
  for approval before execution.

Step by Step
  Executes the installation one step at a time. After each step you
  are asked whether to continue or stop. On failure, you are offered
  the option to roll back all completed steps.
  Useful for testing, troubleshooting, or staged deployments.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OUTPUT / LOG CONSOLE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
The right panel shows real-time output from PowerShell.

Color coding:
  Blue     Informational messages (connecting, installing...)
  Green    Success messages (configured, installed successfully)
  Orange   Warnings (already exists, not configured)
  Red      Errors (failed, exception)
  Gray     Separator lines and log paths

Controls:
  ⏹ Stop          Immediately terminate the running PowerShell process.
                   Use this if an operation is taking too long or you
                   need to cancel mid-execution. Note: steps already
                   completed will not be automatically rolled back.
  A+ / A-          Increase or decrease console font size
  Clear            Clear the console output
  Open Log Folder  Open the folder containing session log files

Log files are saved automatically:
  macOS:    ~/Library/Logs/ETD_Configurator/
  Windows:  %APPDATA%\\ETD_Configurator\\logs\\

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
AUTO-UPDATE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
This application checks for updates automatically at startup.
If a newer version is available, it downloads the updated logic files
in the background and offers to restart the application.
Updates are applied without requiring reinstallation of the .app/.exe.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TROUBLESHOOTING
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PowerShell not found
  Ensure PowerShell 7 (pwsh) is installed.
    macOS:    brew install --cask powershell
    Windows:  winget install Microsoft.PowerShell

Exchange connection fails
  • Verify the Admin UPN is correct
  • Ensure the Exchange Online PowerShell module is installed:
      Install-Module -Name ExchangeOnlineManagement
  • Check that MFA / conditional access is not blocking the connection

Object already exists
  Install skips objects that already exist. Run Verify to see the
  current state. Use Remove to clean up before re-installing.

Permission denied
  The admin account must have the Exchange Administrator role assigned
  in the Microsoft 365 admin center.
"""


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class ETDApp(ctk.CTk):

    # ── Construction ──────────────────────────────────────────────────────

    def __init__(self):
        super().__init__()
        self.title("Cisco ETD  ·  Exchange Online Configurator")
        self.minsize(960, 620)

        # Center window on screen
        w, h = 1180, 740
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        # Cisco brand palette
        self.C_BLUE       = "#049FD9"
        self.C_BLUE_HOVER = "#037EB0"
        self.C_DARK_BLUE  = "#005073"
        self.C_NAVY       = "#1D2B3C"
        self.C_GREEN      = "#6CC04A"
        self.C_GREEN_HOVER= "#549B3A"
        self.C_ORANGE     = "#FF6B00"
        self.C_RED        = "#C0392B"
        self.C_RED_HOVER  = "#922B21"

        self._build_header()
        self._build_body()

    # ── Header ────────────────────────────────────────────────────────────

    def _build_header(self):
        hdr = ctk.CTkFrame(self, corner_radius=0, height=52,
                           fg_color="#005073")
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)

        ctk.CTkLabel(
            hdr,
            text="  Cisco Secure Email Threat Defense  ·  Exchange Online Configurator",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="white",
        ).pack(side="left", padx=20)

        ctk.CTkLabel(
            hdr,
            text=f"v{VERSION}  ",
            font=ctk.CTkFont(size=12),
            text_color="#A0C8D8",
        ).pack(side="right")

        self._update_btn = ctk.CTkButton(
            hdr,
            text="⟳ Check for Updates",
            font=ctk.CTkFont(size=12),
            fg_color="transparent",
            hover_color="#037EB0",
            text_color="#A0C8D8",
            border_width=1,
            border_color="#A0C8D8",
            height=28,
            width=150,
            command=self._manual_update_check,
        )
        self._update_btn.pack(side="right", padx=(0, 10))

    def _manual_update_check(self):
        import threading
        self._update_btn.configure(state="disabled", text="Checking...")

        _done_called = [False]

        def _done(new_version=None, error=None):
            if _done_called[0]:
                return
            _done_called[0] = True
            self._update_btn.configure(state="normal", text="⟳ Check for Updates")
            self.lift()
            self.focus_force()
            if error:
                self._show_dialog("error", "Update Error", f"Could not check for updates:\n{error}")
            elif new_version and new_version > VERSION:
                msg = (
                    f"Version {new_version} is ready.\n\n"
                    "Restart the app to apply the update."
                )
                if self._show_dialog("yesno", "Update Ready", msg):
                    import subprocess, sys
                    self.destroy()
                    subprocess.Popen([sys.executable] + sys.argv)
                    sys.exit(0)
            else:
                self._show_dialog("info", "No Updates", "You are already on the latest version.")

        def _run():
            import socket
            error = None
            new_version = None
            prev_timeout = socket.getdefaulttimeout()
            try:
                socket.setdefaulttimeout(8)
                from updater import check_and_update
                new_version, _ = check_and_update(VERSION)
            except Exception as e:
                error = str(e)
            finally:
                socket.setdefaulttimeout(prev_timeout)
            self.after(0, lambda: _done(new_version, error))

        def _watchdog():
            threading.Event().wait(timeout=15)
            self.after(0, lambda: _done(error="Request timed out."))

        threading.Thread(target=_run, daemon=True).start()
        threading.Thread(target=_watchdog, daemon=True).start()

    def _show_dialog(self, kind, title, message):
        import tkinter as tk
        from tkinter import messagebox
        top = tk.Toplevel(self)
        top.withdraw()
        top.lift()
        top.attributes("-topmost", True)
        if kind == "info":
            messagebox.showinfo(title, message, parent=top)
            result = None
        elif kind == "yesno":
            result = messagebox.askyesno(title, message, parent=top)
        else:
            messagebox.showerror(title, message, parent=top)
            result = None
        top.destroy()
        return result

    # ── Body (form + console) ─────────────────────────────────────────────

    def _build_body(self):
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True)

        # ── Footer ────────────────────────────────────────────────────────
        footer = ctk.CTkFrame(self, corner_radius=0, height=28,
                              fg_color="#D6EAF8")
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        ctk.CTkButton(
            footer, text="About", width=60,
            fg_color="transparent",
            text_color="#005073",
            hover_color="#BEE0F0",
            font=ctk.CTkFont(size=10),
            command=self._show_about,
        ).pack(side="right", padx=10)

        ctk.CTkButton(
            footer, text="Help", width=60,
            fg_color="transparent",
            text_color="#005073",
            hover_color="#BEE0F0",
            font=ctk.CTkFont(size=10),
            command=self._show_help,
        ).pack(side="right", padx=(0, 0))

        # ── Left: form ────────────────────────────────────────────────────
        self.form = ctk.CTkScrollableFrame(
            body, width=420, corner_radius=0,
            fg_color="#EBF5FB",
        )
        self.form.pack(side="left", fill="y")
        self.form.grid_columnconfigure(0, weight=1)

        self._build_form()

        # ── Right: console ────────────────────────────────────────────────
        right = ctk.CTkFrame(body, corner_radius=0, fg_color="transparent")
        right.pack(side="left", fill="both", expand=True)

        console_hdr = ctk.CTkFrame(right, corner_radius=0, height=36,
                                   fg_color="#D6EAF8")
        console_hdr.pack(fill="x")
        console_hdr.pack_propagate(False)

        ctk.CTkLabel(
            console_hdr, text="OUTPUT  /  LOG",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#005073",
        ).pack(side="left", padx=14)

        self.btn_open_log = ctk.CTkButton(
            console_hdr, text="Open Log Folder", width=130,
            fg_color="transparent", border_width=1,
            border_color="#049FD9",
            text_color="#005073",
            hover_color="#BEE0F0",
            font=ctk.CTkFont(size=11),
            command=self._open_log_folder,
        )
        self.btn_open_log.pack(side="right", padx=10, pady=6)

        self.btn_stop = ctk.CTkButton(
            console_hdr, text="⏹ Stop", width=70,
            fg_color="#C0392B", hover_color="#922B21",
            text_color="white",
            font=ctk.CTkFont(size=11, weight="bold"),
            state="disabled",
            command=self._stop_run,
        )
        self.btn_stop.pack(side="right", padx=(0, 4), pady=6)

        ctk.CTkButton(
            console_hdr, text="Clear", width=60,
            fg_color="transparent", border_width=1,
            border_color="#049FD9",
            text_color="#005073",
            hover_color="#BEE0F0",
            font=ctk.CTkFont(size=11),
            command=self._clear_console,
        ).pack(side="right", padx=(0, 4), pady=6)

        # Zoom buttons
        self._console_font_size = 12
        ctk.CTkButton(
            console_hdr, text="A+", width=34,
            fg_color="transparent", border_width=1,
            border_color="#049FD9", text_color="#005073",
            hover_color="#BEE0F0", font=ctk.CTkFont(size=11),
            command=self._zoom_in,
        ).pack(side="right", padx=(0, 2), pady=6)
        ctk.CTkButton(
            console_hdr, text="A-", width=34,
            fg_color="transparent", border_width=1,
            border_color="#049FD9", text_color="#005073",
            hover_color="#BEE0F0", font=ctk.CTkFont(size=11),
            command=self._zoom_out,
        ).pack(side="right", padx=(0, 2), pady=6)

        self.console = ctk.CTkTextbox(
            right,
            font=ctk.CTkFont(family="Courier New", size=self._console_font_size),
            fg_color="#F0F9FF",
            text_color="#1D2B3C",
            wrap="word",
            state="disabled",
        )
        self.console.pack(fill="both", expand=True, padx=10, pady=(6, 4))

        # Configure color tags on the underlying Text widget
        tw = self.console._textbox
        tw.tag_configure("error",   foreground="#C0392B")
        tw.tag_configure("success", foreground="#6CC04A")
        tw.tag_configure("warning", foreground="#FF6B00")
        tw.tag_configure("info",    foreground="#049FD9")
        tw.tag_configure("dim",     foreground="#6B8FA3")

        self._last_log_path = None
        self._stop_fn = None

    # ── Help button helper ─────────────────────────────────────────────────

    def _help_btn(self, parent, key):
        """Return a small 'i' button that opens a help popup for the given key."""
        return ctk.CTkButton(
            parent, text="i", width=22, height=22,
            fg_color="#D6EAF8", hover_color="#BEE0F0",
            text_color="#005073", corner_radius=11,
            font=ctk.CTkFont(size=11, weight="bold"),
            command=lambda k=key: self._show_help_popup(k),
        )

    def _show_help_popup(self, key):
        text = HELP_TEXTS.get(key, "No help available.")
        popup = ctk.CTkToplevel(self)
        popup.title("Help")
        popup.resizable(False, False)
        popup.grab_set()

        frame = ctk.CTkFrame(popup, fg_color="#FFF9C4",
                             border_width=1, border_color="#F9A825",
                             corner_radius=8)
        frame.pack(fill="both", expand=True, padx=6, pady=6)

        ctk.CTkLabel(
            frame, text=text, wraplength=290,
            justify="left", font=ctk.CTkFont(size=11),
            text_color="#1D2B3C",
        ).pack(padx=14, pady=(12, 8), anchor="w")

        ctk.CTkButton(
            frame, text="Close", width=80,
            fg_color="#049FD9", hover_color="#037EB0",
            command=popup.destroy,
        ).pack(pady=(0, 10))

        # Position near mouse pointer but clamped inside the app window
        popup.update_idletasks()
        pw = popup.winfo_reqwidth()
        ph = popup.winfo_reqheight()

        ax = self.winfo_rootx()
        ay = self.winfo_rooty()
        aw = self.winfo_width()
        ah = self.winfo_height()

        mx = self.winfo_pointerx() + 14
        my = self.winfo_pointery() + 14

        x = max(ax + 10, min(mx, ax + aw - pw - 10))
        y = max(ay + 10, min(my, ay + ah - ph - 10))

        popup.geometry(f"{pw}x{ph}+{x}+{y}")

    # ── Form ──────────────────────────────────────────────────────────────

    def _build_form(self):
        F = self.form
        row = [0]

        def r():
            v = row[0]; row[0] += 1; return v

        def section(title):
            fr = ctk.CTkFrame(F, fg_color="transparent", height=28)
            fr.grid(row=r(), column=0, sticky="ew", padx=0, pady=(12, 0))
            ctk.CTkFrame(fr, height=1, fg_color="#049FD9").pack(fill="x", padx=14)
            ctk.CTkLabel(
                fr, text=title,
                font=ctk.CTkFont(size=10, weight="bold"),
                text_color="#005073",
            ).pack(anchor="w", padx=16, pady=(3, 0))

        def lbl_row(text, help_key):
            """A label row with an inline ? help button."""
            fr = ctk.CTkFrame(F, fg_color="transparent")
            fr.grid(row=r(), column=0, sticky="ew", padx=16, pady=(3, 0))
            ctk.CTkLabel(fr, text=text).pack(side="left")
            self._help_btn(fr, help_key).pack(side="left", padx=(6, 0))

        PAD = {"padx": 16, "pady": 3}

        # ── O365 CONNECTION ───────────────────────────────────────────────
        section("O365 CONNECTION")
        lbl_row("Admin UPN:", "admin_upn")
        self.upn_entry = ctk.CTkEntry(F, placeholder_text="admin@tenant.onmicrosoft.com")
        self.upn_entry.grid(row=r(), column=0, sticky="ew", **PAD)

        # ── ETD DEPLOYMENT MODE ───────────────────────────────────────────
        section("ETD DEPLOYMENT MODE")

        mode_hdr = ctk.CTkFrame(F, fg_color="transparent")
        mode_hdr.grid(row=r(), column=0, sticky="w", **PAD)
        self.mode_var = ctk.StringVar(value="Journaling")
        ctk.CTkRadioButton(
            mode_hdr, text="Journaling", variable=self.mode_var,
            value="Journaling", command=self._on_mode_change,
        ).pack(side="left", padx=(0, 24))
        ctk.CTkRadioButton(
            mode_hdr, text="ETD Inline", variable=self.mode_var,
            value="Inline", command=self._on_mode_change,
        ).pack(side="left")
        self._help_btn(mode_hdr, "deployment_mode").pack(side="left", padx=(14, 0))

        # Placeholder row for the dynamic mode panel
        self._mode_panel_row = r()

        # ── ACTIONS ───────────────────────────────────────────────────────
        section("ACTIONS")
        row1 = ctk.CTkFrame(F, fg_color="transparent")
        row1.grid(row=r(), column=0, sticky="ew", **PAD)
        row1.grid_columnconfigure((0, 1), weight=1)

        self.btn_verify = ctk.CTkButton(
            row1, text="Verify",
            fg_color="#049FD9", hover_color="#037EB0",
            command=lambda: self._run("verify"),
        )
        self.btn_verify.grid(row=0, column=0, padx=(0, 4), sticky="ew")

        self.btn_install = ctk.CTkButton(
            row1, text="Install",
            fg_color="#6CC04A", hover_color="#549B3A",
            text_color="white",
            command=lambda: self._run("install"),
        )
        self.btn_install.grid(row=0, column=1, padx=(4, 0), sticky="ew")

        row2 = ctk.CTkFrame(F, fg_color="transparent")
        row2.grid(row=r(), column=0, sticky="ew", **PAD)
        row2.grid_columnconfigure((0, 1), weight=1)

        self.btn_remove = ctk.CTkButton(
            row2, text="Remove",
            fg_color="#C0392B", hover_color="#922B21",
            command=lambda: self._run("remove"),
        )
        self.btn_remove.grid(row=0, column=0, padx=(0, 4), sticky="ew")

        self.btn_export = ctk.CTkButton(
            row2, text="Export .ps1",
            fg_color="#FF6B00", hover_color="#CC5500",
            text_color="white",
            command=self._export,
        )
        self.btn_export.grid(row=0, column=1, padx=(4, 0), sticky="ew")

        # Step-by-Step row
        row3 = ctk.CTkFrame(F, fg_color="transparent")
        row3.grid(row=r(), column=0, sticky="ew", **PAD)
        row3.grid_columnconfigure(0, weight=1)

        self.btn_step = ctk.CTkButton(
            row3, text="Step by Step  ›",
            fg_color="#005073", hover_color="#003D57",
            text_color="white",
            command=self._run_step_by_step,
        )
        self.btn_step.grid(row=0, column=0, padx=(0, 36), sticky="ew")
        self._help_btn(row3, "step_by_step").grid(row=0, column=1, padx=(4, 0))

        self.status_lbl = ctk.CTkLabel(
            F, text="Ready",
            font=ctk.CTkFont(size=11),
            text_color="#005073",
        )
        self.status_lbl.grid(row=r(), column=0, sticky="w", padx=16, pady=(10, 18))

        self._action_buttons = [
            self.btn_verify, self.btn_install,
            self.btn_remove, self.btn_export,
            self.btn_step,
        ]

        # Build mode panels and set initial state
        self._build_journaling_panel()
        self._build_inline_panel()
        self._on_mode_change()

    # ── Journaling panel ──────────────────────────────────────────────────

    def _build_journaling_panel(self):
        F = self.form
        P = {"padx": 12, "pady": 3}

        self.journal_panel = ctk.CTkFrame(F, fg_color="transparent")
        self.journal_panel.grid_columnconfigure(0, weight=1)

        def jlbl(text, help_key, row_idx):
            fr = ctk.CTkFrame(self.journal_panel, fg_color="transparent")
            fr.grid(row=row_idx, column=0, sticky="ew", padx=12, pady=(3, 0))
            ctk.CTkLabel(fr, text=text).pack(side="left")
            self._help_btn(fr, help_key).pack(side="left", padx=(6, 0))

        jlbl("GEO Region:", "geo_region", 0)
        self.journal_geo_var = ctk.StringVar(value=GEO_REGIONS[0])
        ctk.CTkOptionMenu(
            self.journal_panel, variable=self.journal_geo_var, values=GEO_REGIONS,
        ).grid(row=1, column=0, sticky="ew", **P)

        jlbl("Journaling Address:", "journal_address", 2)
        self.journal_entry = ctk.CTkEntry(
            self.journal_panel, placeholder_text="etd-journal@domain.com")
        self.journal_entry.grid(row=3, column=0, sticky="ew", **P)

        jlbl("Notification Alert Email:", "notification_alert", 4)
        self.journal_notif_entry = ctk.CTkEntry(
            self.journal_panel, placeholder_text="alerts@domain.com")
        self.journal_notif_entry.grid(row=5, column=0, sticky="ew", **P)

        seg_row = ctk.CTkFrame(self.journal_panel, fg_color="transparent")
        seg_row.grid(row=6, column=0, sticky="w", **P)
        ctk.CTkLabel(seg_row, text="SEG in front of O365:").pack(side="left", padx=(0, 10))
        self.seg_var = ctk.BooleanVar(value=False)
        ctk.CTkSwitch(
            seg_row, text="", variable=self.seg_var, command=self._on_seg_change,
        ).pack(side="left")
        self._help_btn(seg_row, "seg_in_front").pack(side="left", padx=(10, 0))

        # SEG header field — shown conditionally at row 7/8
        self.seg_header_label = ctk.CTkFrame(self.journal_panel, fg_color="transparent")
        ctk.CTkLabel(self.seg_header_label, text="SEG Header Name:").pack(side="left")
        self._help_btn(self.seg_header_label, "seg_header").pack(side="left", padx=(6, 0))

        self.seg_header_entry = ctk.CTkEntry(
            self.journal_panel, placeholder_text="e.g. X-IronPort-RemoteIP")

    # ── Inline panel ──────────────────────────────────────────────────────

    def _build_inline_panel(self):
        F = self.form
        P = {"padx": 12, "pady": 3}

        self.inline_panel = ctk.CTkFrame(F, fg_color="transparent")
        self.inline_panel.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self.inline_panel, text="SMTP",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#005073",
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(6, 0))

        geo_lbl_fr = ctk.CTkFrame(self.inline_panel, fg_color="transparent")
        geo_lbl_fr.grid(row=1, column=0, sticky="ew", padx=12, pady=(3, 0))
        ctk.CTkLabel(geo_lbl_fr, text="GEO Region:").pack(side="left")
        self._help_btn(geo_lbl_fr, "geo_region").pack(side="left", padx=(6, 0))

        self.inline_geo_var = ctk.StringVar(value=GEO_REGIONS[0])
        ctk.CTkOptionMenu(
            self.inline_panel, variable=self.inline_geo_var, values=GEO_REGIONS,
            command=self._on_geo_change,
        ).grid(row=2, column=0, sticky="ew", **P)

        flows_hdr = ctk.CTkFrame(self.inline_panel, fg_color="transparent")
        flows_hdr.grid(row=3, column=0, sticky="w", padx=12, pady=(8, 0))
        ctk.CTkLabel(
            flows_hdr, text="Active Flows:",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#005073",
        ).pack(side="left")
        self._help_btn(flows_hdr, "active_flows").pack(side="left", padx=(8, 0))

        flows_fr = ctk.CTkFrame(self.inline_panel, fg_color="transparent")
        flows_fr.grid(row=4, column=0, sticky="w", padx=12, pady=(2, 4))
        self.flow_inbound  = ctk.BooleanVar(value=True)
        self.flow_outbound = ctk.BooleanVar(value=True)
        self.flow_internal = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(flows_fr, text="Inbound",  variable=self.flow_inbound,
                        command=self._on_flow_change).pack(side="left", padx=(0, 16))
        ctk.CTkCheckBox(flows_fr, text="Outbound", variable=self.flow_outbound,
                        command=self._on_flow_change).pack(side="left", padx=(0, 16))
        ctk.CTkCheckBox(flows_fr, text="Internal", variable=self.flow_internal,
                        command=self._on_flow_change).pack(side="left")

        # ── Inbound sub-panel (row 5) ──────────────────────────────────────
        self.inbound_panel = ctk.CTkFrame(
            self.inline_panel, corner_radius=8,
            fg_color="#DAEAF6", border_width=1, border_color="#049FD9",
        )
        self.inbound_panel.grid_columnconfigure(0, weight=1)

        ib_hdr = ctk.CTkFrame(self.inbound_panel, fg_color="transparent")
        ib_hdr.grid(row=0, column=0, sticky="w", padx=12, pady=(8, 2))
        ctk.CTkLabel(
            ib_hdr, text="Inbound  —  ETD IP Addresses:",
            font=ctk.CTkFont(size=11, weight="bold"), text_color="#005073",
        ).pack(side="left")
        self._help_btn(ib_hdr, "inbound_ips").pack(side="left", padx=(8, 0))

        self.inbound_ips_text = ctk.CTkTextbox(
            self.inbound_panel, height=90,
            font=ctk.CTkFont(family="Courier New", size=11),
        )
        self.inbound_ips_text.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))

        # ── Outbound sub-panel (row 6) ─────────────────────────────────────
        self.outbound_panel = ctk.CTkFrame(
            self.inline_panel, corner_radius=8,
            fg_color="#DAEAF6", border_width=1, border_color="#049FD9",
        )
        self.outbound_panel.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            self.outbound_panel, text="Outbound:",
            font=ctk.CTkFont(size=11, weight="bold"), text_color="#005073",
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(8, 2))

        sh_lbl = ctk.CTkFrame(self.outbound_panel, fg_color="transparent")
        sh_lbl.grid(row=1, column=0, sticky="w", padx=12)
        ctk.CTkLabel(sh_lbl, text="SmartHost:").pack(side="left")
        self._help_btn(sh_lbl, "smarthost").pack(side="left", padx=(6, 0))

        self.smarthost_entry = ctk.CTkEntry(
            self.outbound_panel, placeholder_text="ob1.hcXXXX.iphmx.com")
        self.smarthost_entry.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 4))

        xp_lbl = ctk.CTkFrame(self.outbound_panel, fg_color="transparent")
        xp_lbl.grid(row=3, column=0, sticky="w", padx=12)
        ctk.CTkLabel(xp_lbl, text="X-CSE-ETD-OUTBOUND-AUTH Value:").pack(side="left")
        self._help_btn(xp_lbl, "xpass").pack(side="left", padx=(6, 0))

        self.xpass_entry = ctk.CTkEntry(self.outbound_panel)
        self.xpass_entry.insert(0, "12345")
        self.xpass_entry.grid(row=4, column=0, sticky="ew", padx=12, pady=(0, 4))

        rule_state_row = ctk.CTkFrame(self.outbound_panel, fg_color="transparent")
        rule_state_row.grid(row=5, column=0, sticky="w", padx=12, pady=(2, 8))
        ctk.CTkLabel(rule_state_row, text="Outbound Tag Rule:").pack(side="left", padx=(0, 10))
        self.outbound_rule_enabled = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(
            rule_state_row, text="Enabled",
            variable=self.outbound_rule_enabled,
        ).pack(side="left")
        self._help_btn(rule_state_row, "outbound_rule").pack(side="left", padx=(10, 0))

        # ── Internal sub-panel (row 7) ─────────────────────────────────────
        self.internal_panel = ctk.CTkFrame(
            self.inline_panel, corner_radius=8,
            fg_color="#DAEAF6", border_width=1, border_color="#049FD9",
        )
        self.internal_panel.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            self.internal_panel, text="Internal:",
            font=ctk.CTkFont(size=11, weight="bold"), text_color="#005073",
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(8, 2))

        ij_lbl = ctk.CTkFrame(self.internal_panel, fg_color="transparent")
        ij_lbl.grid(row=1, column=0, sticky="w", padx=12)
        ctk.CTkLabel(ij_lbl, text="Journaling Address:").pack(side="left")
        self._help_btn(ij_lbl, "internal_journal").pack(side="left", padx=(6, 0))

        self.internal_journal_entry = ctk.CTkEntry(
            self.internal_panel, placeholder_text="etd-journal@domain.com")
        self.internal_journal_entry.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 4))

        in_lbl = ctk.CTkFrame(self.internal_panel, fg_color="transparent")
        in_lbl.grid(row=3, column=0, sticky="w", padx=12)
        ctk.CTkLabel(in_lbl, text="Notification Alert Email:").pack(side="left")
        self._help_btn(in_lbl, "notification_alert").pack(side="left", padx=(6, 0))

        self.internal_notif_entry = ctk.CTkEntry(
            self.internal_panel, placeholder_text="alerts@domain.com")
        self.internal_notif_entry.grid(row=4, column=0, sticky="ew", padx=12, pady=(0, 8))

    # ── Visibility toggles ────────────────────────────────────────────────

    def _on_mode_change(self):
        if self.mode_var.get() == "Journaling":
            self.journal_panel.grid(
                row=self._mode_panel_row, column=0, sticky="ew", padx=4, pady=4)
            self.inline_panel.grid_remove()
            self._on_seg_change()
            w, h = 1180, 740
        else:
            self.inline_panel.grid(
                row=self._mode_panel_row, column=0, sticky="ew", padx=4, pady=4)
            self.journal_panel.grid_remove()
            self._on_flow_change()
            w, h = 1180, 1000
        x = self.winfo_x()
        y = self.winfo_y()
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _on_flow_change(self):
        if self.flow_inbound.get():
            self.inbound_panel.grid(row=5, column=0, sticky="ew", padx=12, pady=(4, 2))
            self._populate_inbound_ips()
        else:
            self.inbound_panel.grid_remove()
        if self.flow_outbound.get():
            self.outbound_panel.grid(row=6, column=0, sticky="ew", padx=12, pady=2)
            self._populate_smarthost()
        else:
            self.outbound_panel.grid_remove()
        if self.flow_internal.get():
            self.internal_panel.grid(row=7, column=0, sticky="ew", padx=12, pady=(2, 4))
        else:
            self.internal_panel.grid_remove()

    def _on_seg_change(self):
        if self.seg_var.get():
            self.seg_header_label.grid(
                row=7, column=0, sticky="ew", padx=12, pady=(6, 0))
            self.seg_header_entry.grid(
                row=8, column=0, sticky="ew", padx=12, pady=(0, 8))
        else:
            self.seg_header_label.grid_remove()
            self.seg_header_entry.grid_remove()

    def _on_geo_change(self, *_):
        self._populate_inbound_ips()
        self._populate_smarthost()

    def _populate_inbound_ips(self):
        geo = self.inline_geo_var.get()
        ips = GEO_IPS.get(geo, [])
        self.inbound_ips_text.delete("1.0", "end")
        if ips:
            self.inbound_ips_text.insert("1.0", "\n".join(ips))

    def _populate_smarthost(self):
        geo  = self.inline_geo_var.get()
        host = GEO_OUTBOUND_HOSTS.get(geo, "")
        self.smarthost_entry.delete(0, "end")
        if host:
            self.smarthost_entry.insert(0, host)

    # ── Config collection + validation ────────────────────────────────────

    def _collect(self, operation):
        admin_upn = self.upn_entry.get().strip()
        mode      = self.mode_var.get()

        if not _valid_email(admin_upn):
            messagebox.showerror("Validation", "Invalid Admin UPN format.")
            return None

        if mode == "Journaling":
            geo          = self.journal_geo_var.get()
            journal      = self.journal_entry.get().strip()
            notification = self.journal_notif_entry.get().strip()

            if not _valid_email(journal):
                messagebox.showerror("Validation", "Invalid Journaling Address format.")
                return None

            seg_in_front    = self.seg_var.get()
            seg_header_name = self.seg_header_entry.get().strip() if seg_in_front else ""

            return {
                "admin_upn":          admin_upn,
                "journal_address":    journal,
                "notification_alert": notification,
                "deployment_mode":    mode,
                "geo":                geo,
                "etd_ips":            [],
                "seg_in_front":       seg_in_front,
                "seg_header_name":    seg_header_name,
                "seg_ips":            [],
                "flows":              {},
                "smart_host":         "",
                "xpass_value":        "",
                "operation":          operation,
            }

        else:  # Inline
            geo   = self.inline_geo_var.get()
            flows = {
                "inbound":  self.flow_inbound.get(),
                "outbound": self.flow_outbound.get(),
                "internal": self.flow_internal.get(),
            }

            if not any(flows.values()):
                messagebox.showerror("Validation", "Select at least one flow.")
                return None

            etd_ips    = []
            smart_host = ""
            xpass      = "12345"
            journal    = ""
            notification = ""
            outbound_rule_enabled = True

            if flows["inbound"]:
                etd_ips = _parse_ips(self.inbound_ips_text.get("1.0", "end"))
                if not etd_ips:
                    messagebox.showerror("Validation", "Enter ETD IP addresses for Inbound.")
                    return None

            if flows["outbound"]:
                smart_host            = self.smarthost_entry.get().strip()
                xpass                 = self.xpass_entry.get().strip() or "12345"
                outbound_rule_enabled = self.outbound_rule_enabled.get()
                if not smart_host:
                    messagebox.showerror("Validation", "SmartHost is required for Outbound.")
                    return None

            if flows["internal"]:
                journal      = self.internal_journal_entry.get().strip()
                notification = self.internal_notif_entry.get().strip()
                if not _valid_email(journal):
                    messagebox.showerror("Validation", "Invalid Journaling Address for Internal flow.")
                    return None

            return {
                "admin_upn":           admin_upn,
                "journal_address":     journal,
                "notification_alert":  notification,
                "deployment_mode":     mode,
                "flows":               flows,
                "geo":                 geo,
                "etd_ips":             etd_ips,
                "smart_host":          smart_host,
                "outbound_rule_enabled": outbound_rule_enabled,
                "seg_in_front":        False,
                "seg_ips":             [],
                "xpass_value":         xpass,
                "operation":           operation,
            }

    # ── Run ───────────────────────────────────────────────────────────────

    def _run(self, operation):
        config = self._collect(operation)
        if config is None:
            return

        if (operation == "install"
                and config.get("deployment_mode") == "Inline"
                and config.get("flows", {}).get("outbound")
                and config.get("outbound_rule_enabled")):
            if not self._show_dialog(
                "yesno",
                "Outbound Tag Rule — Warning",
                "The Outbound Tag Rule is currently set to Enabled.\n\n"
                "This rule takes effect immediately and will affect live mail flow "
                "as soon as it is installed.\n\n"
                "Recommendation: install with the rule Disabled and enable it "
                "manually once you have verified the connector is working correctly.\n\n"
                "Do you want to continue with the rule Enabled?",
            ):
                return

        script = generate_script(config)

        self._set_btns("disabled")
        self._status(f"Running {operation}...", "gray")
        self._log(f"{'=' * 56}", "dim")
        self._log(
            f"  {operation.upper()}  |  {config['deployment_mode']}  |  {config['geo']}",
            "info",
        )
        self._log(f"{'=' * 56}", "dim")

        q = queue.Queue()

        def on_output(line):
            q.put(("line", line))

        def on_done(success, log_path):
            q.put(("done", success, log_path))

        def _poll():
            while True:
                try:
                    item = q.get_nowait()
                except queue.Empty:
                    break

                if item[0] == "line":
                    self._log_auto(item[1])

                elif item[0] == "done":
                    _, success, log_path = item
                    self._last_log_path = log_path
                    while True:
                        try:
                            extra = q.get_nowait()
                            if extra[0] == "line":
                                self._log_auto(extra[1])
                        except queue.Empty:
                            break
                    self._log(f"{'=' * 56}", "dim")
                    if success:
                        self._log("  Operation completed successfully.", "success")
                        self._status(f"Last: {operation} OK", "green")
                    else:
                        self._log("  Operation completed with errors.", "error")
                        self._status(f"Last: {operation} FAILED", "#E74C3C")
                    self._log(f"  Log saved: {log_path}", "dim")
                    self._set_btns("normal")

                    if not success and operation == "install":
                        if messagebox.askyesno(
                            "Installation Error",
                            "Errors occurred during installation.\n\n"
                            "Do you want to undo the changes made?",
                        ):
                            self._run("remove")
                    return

            self.after(50, _poll)

        self.after(50, _poll)
        self._stop_fn = run_script(script, on_output, on_done)

    # ── Step-by-Step ──────────────────────────────────────────────────────

    def _run_step_by_step(self):
        config = self._collect("install")
        if config is None:
            return

        steps = generate_steps(config)
        if not steps:
            messagebox.showinfo("Step by Step", "No steps to execute for this configuration.")
            return

        self._set_btns("disabled")
        self._log(f"{'=' * 56}", "dim")
        self._log(
            f"  STEP-BY-STEP  |  {config['deployment_mode']}  |  {config['geo']}",
            "info",
        )
        self._log(f"  Total steps: {len(steps)}", "info")
        self._log(f"{'=' * 56}", "dim")
        self._run_next_step(steps, 0, config)

    def _run_next_step(self, steps, idx, config):
        step_name, script = steps[idx]
        total = len(steps)

        self._log(f"  Step {idx + 1}/{total}: {step_name}", "info")
        self._status(f"Step {idx + 1}/{total}: {step_name}", "#049FD9")

        q = queue.Queue()

        def on_output(line):
            q.put(("line", line))

        def on_done(success, log_path):
            q.put(("done", success, log_path))

        def _poll():
            while True:
                try:
                    item = q.get_nowait()
                except queue.Empty:
                    break

                if item[0] == "line":
                    self._log_auto(item[1])

                elif item[0] == "done":
                    _, success, log_path = item
                    self._last_log_path = log_path
                    while True:
                        try:
                            extra = q.get_nowait()
                            if extra[0] == "line":
                                self._log_auto(extra[1])
                        except queue.Empty:
                            break

                    if not success:
                        self._log(f"  Step '{step_name}' FAILED.", "error")
                        self._status(f"Step {idx + 1} FAILED", "#E74C3C")
                        self._set_btns("normal")
                        if messagebox.askyesno(
                            "Step Failed",
                            f"Step '{step_name}' failed.\n\n"
                            "Do you want to undo all steps completed so far?",
                        ):
                            self._run("remove")
                        return

                    self._log(f"  Step '{step_name}' completed successfully.", "success")

                    if idx + 1 < total:
                        next_name = steps[idx + 1][0]
                        if messagebox.askyesno(
                            "Continue?",
                            f"Step {idx + 1}/{total}  '{step_name}'  completed.\n\n"
                            f"Continue with step {idx + 2}/{total}:\n"
                            f"  '{next_name}' ?",
                        ):
                            self._run_next_step(steps, idx + 1, config)
                        else:
                            self._log("Step-by-step stopped by user.", "warning")
                            self._status(f"Stopped at step {idx + 1}/{total}", "#FF6B00")
                            self._set_btns("normal")
                    else:
                        self._log(f"{'=' * 56}", "dim")
                        self._log("  All steps completed successfully.", "success")
                        self._status("Step-by-step: all done", "green")
                        self._set_btns("normal")
                    return

            self.after(50, _poll)

        self.after(50, _poll)
        self._stop_fn = run_script(script, on_output, on_done)

    # ── Export ────────────────────────────────────────────────────────────

    def _export(self):
        config = self._collect("install")
        if config is None:
            return

        script = generate_script(config)
        path = filedialog.asksaveasfilename(
            defaultextension=".ps1",
            filetypes=[("PowerShell Script", "*.ps1"), ("All Files", "*.*")],
            initialfile=f"ETD_{config['deployment_mode'].lower()}_install.ps1",
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(script)
            self._log(f"Script exported to: {path}", "success")
            messagebox.showinfo("Export", f"Script saved:\n{path}")

    # ── Console helpers ───────────────────────────────────────────────────

    def _log(self, text, tag=None):
        self.console.configure(state="normal")
        tw = self.console._textbox
        if tag:
            tw.insert("end", text + "\n", tag)
        else:
            self.console.insert("end", text + "\n")
        self.console.see("end")
        self.console.configure(state="disabled")

    def _log_auto(self, text):
        """Log with automatic color detection based on content."""
        lo = text.lower()
        if "error" in lo or "failed" in lo or "exception" in lo:
            tag = "error"
        elif "successfully" in lo or "configured" in lo or "completed" in lo:
            tag = "success"
        elif "not configured" in lo or "already exists" in lo or "warning" in lo:
            tag = "warning"
        elif text.startswith("=") or text.startswith("["):
            tag = "dim"
        elif "installing" in lo or "removing" in lo or "connecting" in lo:
            tag = "info"
        else:
            tag = None
        self._log(text, tag)

    def _clear_console(self):
        self.console.configure(state="normal")
        self.console.delete("1.0", "end")
        self.console.configure(state="disabled")

    def _open_log_folder(self):
        from executor import _log_dir
        folder = str(_log_dir())
        if sys.platform == "win32":
            os.startfile(folder)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", folder])
        else:
            subprocess.Popen(["xdg-open", folder])

    # ── Zoom ──────────────────────────────────────────────────────────────

    def _zoom_in(self):
        if self._console_font_size < 28:
            self._console_font_size += 2
            self.console.configure(
                font=ctk.CTkFont(family="Courier New", size=self._console_font_size))

    def _zoom_out(self):
        if self._console_font_size > 8:
            self._console_font_size -= 2
            self.console.configure(
                font=ctk.CTkFont(family="Courier New", size=self._console_font_size))

    # ── About ─────────────────────────────────────────────────────────────

    def _show_about(self):
        win = ctk.CTkToplevel(self)
        win.title("About")
        win.geometry("420x220")
        win.resizable(False, False)
        win.grab_set()

        ctk.CTkLabel(
            win,
            text="Cisco ETD  ·  Exchange Online Configurator",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#005073",
        ).pack(pady=(24, 2))

        ctk.CTkLabel(
            win,
            text=f"Version {VERSION}  ·  {VERSION_DATE}",
            font=ctk.CTkFont(size=10),
            text_color="#6B8FA3",
        ).pack(pady=(0, 8))

        ctk.CTkLabel(
            win,
            text=(
                "This tool is intended for internal use only and was developed\n"
                "to simplify the configuration of Exchange Online with\n"
                "Cisco Secure Email Threat Defense.\n\n"
                "It is provided without official support.\n"
                "Use it at your own risk."
            ),
            font=ctk.CTkFont(size=11),
            text_color="#1D2B3C",
            justify="center",
        ).pack(padx=20)

        ctk.CTkButton(
            win, text="Close", width=100,
            fg_color="#049FD9", hover_color="#037EB0",
            command=win.destroy,
        ).pack(pady=(16, 0))

    # ── Help manual ───────────────────────────────────────────────────────

    def _show_help(self):
        win = ctk.CTkToplevel(self)
        win.title("Help  —  ETD Configurator User Manual")
        win.resizable(True, True)

        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w, h = min(760, sw - 100), min(700, sh - 100)
        x = (sw - w) // 2
        y = (sh - h) // 2
        win.geometry(f"{w}x{h}+{x}+{y}")
        win.grab_set()

        # Header
        hdr = ctk.CTkFrame(win, corner_radius=0, height=44, fg_color="#005073")
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(
            hdr, text="  User Manual  —  ETD Configurator",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="white",
        ).pack(side="left", padx=16)

        # Scrollable text area
        txt = ctk.CTkTextbox(
            win,
            font=ctk.CTkFont(family="Courier New", size=12),
            fg_color="#F8FBFF",
            text_color="#1D2B3C",
            wrap="word",
        )
        txt.pack(fill="both", expand=True, padx=10, pady=(8, 4))

        content = USER_MANUAL.format(version=VERSION, date=VERSION_DATE)
        txt.insert("1.0", content)
        txt.configure(state="disabled")

        ctk.CTkButton(
            win, text="Close", width=100,
            fg_color="#049FD9", hover_color="#037EB0",
            command=win.destroy,
        ).pack(pady=(0, 10))

    # ── Status + button helpers ───────────────────────────────────────────

    def _status(self, text, color):
        self.status_lbl.configure(text=text, text_color=color)

    def _set_btns(self, state):
        for btn in self._action_buttons:
            btn.configure(state=state)
        self.btn_stop.configure(state="normal" if state == "disabled" else "disabled")

    def _stop_run(self):
        if self._stop_fn:
            self._stop_fn()
            self._stop_fn = None
        self._log("  Stopping...", "warning")
        self.btn_stop.configure(state="disabled")
