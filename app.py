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
from ps_generator import generate_script
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
        self.C_BLUE       = "#049FD9"   # Cisco primary blue
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

        PAD = {"padx": 16, "pady": 3}

        # ── O365 CONNECTION ───────────────────────────────────────────────
        section("O365 CONNECTION")
        ctk.CTkLabel(F, text="Admin UPN:").grid(row=r(), column=0, sticky="w", **PAD)
        self.upn_entry = ctk.CTkEntry(F, placeholder_text="admin@tenant.onmicrosoft.com")
        self.upn_entry.grid(row=r(), column=0, sticky="ew", **PAD)

        # ── ETD DEPLOYMENT MODE ───────────────────────────────────────────
        section("ETD DEPLOYMENT MODE")
        self.mode_var = ctk.StringVar(value="Journaling")
        mode_row = ctk.CTkFrame(F, fg_color="transparent")
        mode_row.grid(row=r(), column=0, sticky="w", **PAD)
        ctk.CTkRadioButton(
            mode_row, text="Journaling", variable=self.mode_var,
            value="Journaling", command=self._on_mode_change,
        ).pack(side="left", padx=(0, 24))
        ctk.CTkRadioButton(
            mode_row, text="ETD Inline", variable=self.mode_var,
            value="Inline", command=self._on_mode_change,
        ).pack(side="left")

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

        self.status_lbl = ctk.CTkLabel(
            F, text="Ready",
            font=ctk.CTkFont(size=11),
            text_color="#005073",
        )
        self.status_lbl.grid(row=r(), column=0, sticky="w", padx=16, pady=(10, 18))

        self._action_buttons = [
            self.btn_verify, self.btn_install,
            self.btn_remove, self.btn_export,
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

        ctk.CTkLabel(self.journal_panel, text="GEO Region:").grid(
            row=0, column=0, sticky="w", **P)
        self.journal_geo_var = ctk.StringVar(value=GEO_REGIONS[0])
        ctk.CTkOptionMenu(
            self.journal_panel, variable=self.journal_geo_var, values=GEO_REGIONS,
        ).grid(row=1, column=0, sticky="ew", **P)

        ctk.CTkLabel(self.journal_panel, text="Journaling Address:").grid(
            row=2, column=0, sticky="w", **P)
        self.journal_entry = ctk.CTkEntry(
            self.journal_panel, placeholder_text="etd-journal@domain.com")
        self.journal_entry.grid(row=3, column=0, sticky="ew", **P)

        ctk.CTkLabel(self.journal_panel, text="Notification Alert Email:").grid(
            row=4, column=0, sticky="w", **P)
        self.journal_notif_entry = ctk.CTkEntry(
            self.journal_panel, placeholder_text="alerts@domain.com")
        self.journal_notif_entry.grid(row=5, column=0, sticky="ew", **P)

        seg_row = ctk.CTkFrame(self.journal_panel, fg_color="transparent")
        seg_row.grid(row=6, column=0, sticky="w", **P)
        ctk.CTkLabel(seg_row, text="SEG in front of O365:").pack(side="left", padx=(0, 12))
        self.seg_var = ctk.BooleanVar(value=False)
        ctk.CTkSwitch(
            seg_row, text="", variable=self.seg_var, command=self._on_seg_change,
        ).pack(side="left")

        # SEG header field — shown conditionally at row 7
        self.seg_header_label = ctk.CTkLabel(
            self.journal_panel, text="SEG Header Name:")
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

        ctk.CTkLabel(self.inline_panel, text="GEO Region:").grid(
            row=1, column=0, sticky="w", **P)
        self.inline_geo_var = ctk.StringVar(value=GEO_REGIONS[0])
        ctk.CTkOptionMenu(
            self.inline_panel, variable=self.inline_geo_var, values=GEO_REGIONS,
            command=self._on_geo_change,
        ).grid(row=2, column=0, sticky="ew", **P)

        ctk.CTkLabel(
            self.inline_panel, text="Active Flows:",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#005073",
        ).grid(row=3, column=0, sticky="w", padx=12, pady=(8, 2))

        flows_fr = ctk.CTkFrame(self.inline_panel, fg_color="transparent")
        flows_fr.grid(row=4, column=0, sticky="w", padx=12, pady=(0, 4))
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
        ctk.CTkLabel(
            self.inbound_panel, text="Inbound  —  ETD IP Addresses:",
            font=ctk.CTkFont(size=11, weight="bold"), text_color="#005073",
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(8, 2))
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
        ctk.CTkLabel(self.outbound_panel, text="SmartHost:").grid(
            row=1, column=0, sticky="w", padx=12)
        self.smarthost_entry = ctk.CTkEntry(
            self.outbound_panel, placeholder_text="ob1.hcXXXX.iphmx.com")
        self.smarthost_entry.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 4))
        ctk.CTkLabel(
            self.outbound_panel, text="X-CSE-ETD-OUTBOUND-AUTH Value:").grid(
            row=3, column=0, sticky="w", padx=12)
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
        ctk.CTkLabel(self.internal_panel, text="Journaling Address:").grid(
            row=1, column=0, sticky="w", padx=12)
        self.internal_journal_entry = ctk.CTkEntry(
            self.internal_panel, placeholder_text="etd-journal@domain.com")
        self.internal_journal_entry.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 4))
        ctk.CTkLabel(self.internal_panel, text="Notification Alert Email:").grid(
            row=3, column=0, sticky="w", padx=12)
        self.internal_notif_entry = ctk.CTkEntry(
            self.internal_panel, placeholder_text="alerts@domain.com")
        self.internal_notif_entry.grid(row=4, column=0, sticky="ew", padx=12, pady=(0, 8))

    # ── Visibility toggles ────────────────────────────────────────────────

    def _on_mode_change(self):
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
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
        x = (sw - w) // 2
        y = (sh - h) // 2
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
                row=7, column=0, sticky="w", padx=12, pady=(6, 0))
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
                "admin_upn":         admin_upn,
                "journal_address":   journal,
                "notification_alert": notification,
                "deployment_mode":   mode,
                "geo":               geo,
                "etd_ips":           [],
                "seg_in_front":      seg_in_front,
                "seg_header_name":   seg_header_name,
                "seg_ips":           [],
                "flows":             {},
                "smart_host":        "",
                "xpass_value":       "",
                "operation":         operation,
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

            if flows["inbound"]:
                etd_ips = _parse_ips(self.inbound_ips_text.get("1.0", "end"))
                if not etd_ips:
                    messagebox.showerror("Validation", "Enter ETD IP addresses for Inbound.")
                    return None

            if flows["outbound"]:
                smart_host           = self.smarthost_entry.get().strip()
                xpass                = self.xpass_entry.get().strip() or "12345"
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
                "admin_upn":          admin_upn,
                "journal_address":    journal,
                "notification_alert": notification,
                "deployment_mode":    mode,
                "flows":              flows,
                "geo":                geo,
                "etd_ips":            etd_ips,
                "smart_host":              smart_host,
                "outbound_rule_enabled":   outbound_rule_enabled if flows["outbound"] else True,
                "seg_in_front":            False,
                "seg_ips":            [],
                "xpass_value":        xpass,
                "operation":          operation,
            }

    # ── Run ───────────────────────────────────────────────────────────────

    def _run(self, operation):
        config = self._collect(operation)
        if config is None:
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

        # El hilo de fondo NUNCA toca la UI — solo escribe en la Queue.
        # El hilo principal vacía la Queue cada 50 ms con _poll.
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
                    # Vaciar líneas que llegaron junto al "done"
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

                    # Si install falló, ofrecer rollback
                    if not success and operation == "install":
                        if messagebox.askyesno(
                            "Error durante la instalación",
                            "Se produjeron errores durante la instalación.\n\n"
                            "¿Deseas deshacer los cambios realizados?",
                        ):
                            self._run("remove")

                    return  # dejar de hacer poll

            self.after(50, _poll)

        self.after(50, _poll)
        run_script(script, on_output, on_done)

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

    # ── Status + button helpers ───────────────────────────────────────────

    def _status(self, text, color):
        self.status_lbl.configure(text=text, text_color=color)

    def _set_btns(self, state):
        for btn in self._action_buttons:
            btn.configure(state=state)
