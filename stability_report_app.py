"""
Stability Report Generator
--------------------------
GUI launcher for stability_report_v2.py.

Point at a study folder (e.g. HE4_Stab11/) containing TP subfolders.
The tool generates:
  - One per-timepoint Excel report per TP folder
  - A combined trends Excel report with charts
  - Three Tableau-ready CSV files

To rebuild the .exe after updating stability_report_v2.py:
    Windows:
    pyinstaller --onefile --windowed --name "StabilityReportGenerator" ^
        --add-data "stability_report_v2.py;." stability_report_app.py
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import traceback
import os
import re
import sys

# ── Ensure core script is importable whether running as .py or .exe ──
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, BASE_DIR)

try:
    from stability_report_v2 import run_study
except ImportError as e:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Missing File",
        f"Could not find stability_report_v2.py.\n\n"
        f"Make sure stability_report_v2.py is in the same folder "
        f"as this application.\n\nDetails: {e}"
    )
    sys.exit(1)


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def discover_timepoints(study_folder):
    """Return sorted list of TP subfolder names found in study_folder."""
    tps = []
    for item in os.listdir(study_folder):
        if os.path.isdir(os.path.join(study_folder, item)) and \
                re.match(r'^TP\d+$', item, re.IGNORECASE):
            tps.append(item.upper())
    return sorted(tps, key=lambda t: int(re.sub(r'[^0-9]', '', t)))


def count_raven_files(study_folder, tp_list):
    """Count .xlsx files across all TP folders."""
    total = 0
    for tp in tp_list:
        tp_path = os.path.join(study_folder, tp)
        if os.path.isdir(tp_path):
            total += sum(1 for f in os.listdir(tp_path)
                         if f.endswith(".xlsx") and not f.startswith("~"))
    return total


# ─────────────────────────────────────────────
# MAIN APPLICATION
# ─────────────────────────────────────────────

class StabilityReportApp:

    # Colours
    BG_DARK   = "#1F3864"
    BG_MID    = "#2E75B6"
    BG_WHITE  = "#FFFFFF"
    BG_CARD   = "#F7F9FC"
    FG_WHITE  = "#FFFFFF"
    FG_DARK   = "#1F3864"
    FG_GREY   = "#555555"
    FG_LIGHT  = "#BDD7EE"
    BTN_GREEN = "#C6EFCE"
    BTN_GFONT = "#006100"
    BTN_BLUE  = "#2E75B6"
    ERR_RED   = "#9C0006"

    def __init__(self, root):
        self.root = root
        self.root.title("Stability Report Generator")
        self.root.resizable(False, True)
        self.root.configure(bg=self.BG_DARK)

        win_w, win_h = 620, 580
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        root.geometry(f"{win_w}x{win_h}+{(sw-win_w)//2}+{(sh-win_h)//2}")

        self.study_folder_var = tk.StringVar()
        self.study_folder_var.trace_add("write", self._on_folder_change)

        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self.root, bg=self.BG_DARK, pady=18)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Stability Report Generator",
                 font=("Arial", 18, "bold"),
                 fg=self.FG_WHITE, bg=self.BG_DARK).pack()
        tk.Label(hdr, text="Fujirebio Diagnostics — Assay Stability Studies",
                 font=("Arial", 10),
                 fg=self.FG_LIGHT, bg=self.BG_DARK).pack()

        # Card
        card = tk.Frame(self.root, bg=self.BG_CARD,
                        padx=24, pady=18, relief="flat")
        card.pack(fill="both", expand=True, padx=16, pady=(0, 0))
        card.columnconfigure(0, weight=1)

        # ── Study folder ──
        tk.Label(card, text="Study Folder",
                 font=("Arial", 10, "bold"),
                 fg=self.FG_DARK, bg=self.BG_CARD,
                 anchor="w").grid(row=0, column=0, sticky="w", pady=(0, 2))
        tk.Label(card,
                 text="Select the study folder containing TP1, TP2, … subfolders  "
                      "(e.g. HE4_Stab11/)",
                 font=("Arial", 9), fg=self.FG_GREY,
                 bg=self.BG_CARD, anchor="w"
                 ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 5))

        folder_row = tk.Frame(card, bg=self.BG_CARD)
        folder_row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        folder_row.columnconfigure(0, weight=1)

        self.folder_entry = tk.Entry(folder_row,
                                     textvariable=self.study_folder_var,
                                     font=("Arial", 10), relief="solid", bd=1)
        self.folder_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        tk.Button(folder_row, text="Browse…", font=("Arial", 10),
                  bg=self.BTN_BLUE, fg=self.FG_WHITE, relief="flat",
                  padx=12, cursor="hand2",
                  command=self._browse_study_folder
                  ).grid(row=0, column=1)

        # ── Discovery panel ──
        disc_frame = tk.LabelFrame(card, text="Study Contents",
                                   font=("Arial", 9, "bold"),
                                   fg=self.FG_DARK, bg=self.BG_CARD,
                                   padx=10, pady=8)
        disc_frame.grid(row=3, column=0, columnspan=2,
                        sticky="ew", pady=(8, 0))
        disc_frame.columnconfigure(1, weight=1)

        labels = ["Study name:", "Timepoints found:", "Raven files found:"]
        self.disc_vars = [tk.StringVar(value="—") for _ in labels]
        for i, (lbl, var) in enumerate(zip(labels, self.disc_vars)):
            tk.Label(disc_frame, text=lbl, font=("Arial", 9, "bold"),
                     fg=self.FG_DARK, bg=self.BG_CARD,
                     anchor="w").grid(row=i, column=0, sticky="w", pady=1)
            tk.Label(disc_frame, textvariable=var,
                     font=("Arial", 9), fg=self.FG_GREY,
                     bg=self.BG_CARD, anchor="w"
                     ).grid(row=i, column=1, sticky="w", padx=(8, 0), pady=1)

        # ── Outputs panel ──
        out_frame = tk.LabelFrame(card, text="Outputs  (saved into the study folder)",
                                  font=("Arial", 9, "bold"),
                                  fg=self.FG_DARK, bg=self.BG_CARD,
                                  padx=10, pady=8)
        out_frame.grid(row=4, column=0, columnspan=2,
                       sticky="ew", pady=(10, 0))

        outputs = [
            "✔  Per-timepoint Excel report   (one per TP)",
            "✔  Combined trends Excel report  (with charts)",
            "✔  Tableau CSV — Run Summary",
            "✔  Tableau CSV — Grand Means",
            "✔  Tableau CSV — Replicates",
        ]
        for txt in outputs:
            tk.Label(out_frame, text=txt,
                     font=("Arial", 9), fg="#444",
                     bg=self.BG_CARD, anchor="w"
                     ).pack(anchor="w")

        # ── Status + progress ──
        self.status_var = tk.StringVar(value="Select a study folder to begin.")
        self.status_lbl = tk.Label(card, textvariable=self.status_var,
                                   font=("Arial", 9, "italic"),
                                   fg=self.FG_GREY, bg=self.BG_CARD,
                                   anchor="w")
        self.status_lbl.grid(row=5, column=0, columnspan=2,
                             sticky="w", pady=(12, 2))

        self.progress = ttk.Progressbar(card, mode="indeterminate",
                                        length=540)
        self.progress.grid(row=6, column=0, columnspan=2,
                           sticky="ew", pady=(0, 0))

        # ── Run button ──
        btn_frame = tk.Frame(self.root, bg=self.BG_DARK, pady=12)
        btn_frame.pack(fill="x", padx=16)

        self.run_btn = tk.Button(
            btn_frame,
            text="Generate All Reports",
            font=("Arial", 12, "bold"),
            bg=self.BTN_GREEN, fg=self.BTN_GFONT,
            relief="flat", padx=20, pady=10,
            cursor="hand2",
            command=self._run
        )
        self.run_btn.pack(fill="x")

    # ── Folder interaction ────────────────────────────────────────────

    def _browse_study_folder(self):
        folder = filedialog.askdirectory(title="Select Study Folder")
        if folder:
            self.study_folder_var.set(folder)

    def _on_folder_change(self, *_):
        folder = self.study_folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            for v in self.disc_vars:
                v.set("—")
            self._set_status("Select a study folder to begin.", self.FG_GREY)
            return

        # Auto-discover contents
        study_name = re.sub(r'(?i)(Stab)(\d+)', r'\1 \2',
                            os.path.basename(folder.rstrip("/\\"))
                            .replace("_", " "))
        tps   = discover_timepoints(folder)
        n_files = count_raven_files(folder, tps)

        self.disc_vars[0].set(study_name)
        self.disc_vars[1].set(", ".join(tps) if tps else "None found")
        self.disc_vars[2].set(str(n_files) if n_files else "0")

        if not tps:
            self._set_status(
                "⚠ No TP subfolders found. Check folder structure.",
                self.ERR_RED)
        elif n_files == 0:
            self._set_status(
                "⚠ TP folders found but no .xlsx files inside.",
                self.ERR_RED)
        else:
            self._set_status(
                f"Ready — {len(tps)} timepoint(s), {n_files} file(s) found.",
                "#006100")

    # ── Report generation ─────────────────────────────────────────────

    def _run(self):
        folder = self.study_folder_var.get().strip()

        if not folder:
            messagebox.showwarning("No Folder Selected",
                                   "Please select a study folder first.")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("Invalid Folder",
                                 f"Folder does not exist:\n{folder}")
            return

        tps = discover_timepoints(folder)
        if not tps:
            messagebox.showerror(
                "No Timepoints Found",
                f"No TP subfolders (TP1, TP2, …) were found in:\n{folder}\n\n"
                "Please check your folder structure matches the instructions."
            )
            return

        n_files = count_raven_files(folder, tps)
        if n_files == 0:
            messagebox.showerror(
                "No Raven Files Found",
                "TP folders were found but contain no .xlsx files.\n\n"
                "Make sure your renamed Raven output files are inside "
                "the correct TP subfolder."
            )
            return

        # Disable UI and start
        self.run_btn.config(state="disabled", bg="#CCCCCC", fg="#888888")
        self.progress.start(10)
        self._set_status("Starting…", self.BTN_BLUE)

        thread = threading.Thread(
            target=self._generate, args=(folder, tps), daemon=True
        )
        thread.start()

    def _generate(self, folder, tps):
        try:
            # Live status updates via callback
            original_print = __builtins__.__dict__.get('print', print) \
                if hasattr(__builtins__, '__dict__') else print

            completed_tps  = []
            total_records  = 0
            total_reps     = 0

            import io, contextlib

            # Intercept stdout to update status live during processing
            class StatusCapture(io.StringIO):
                def __init__(self, app, tps_count):
                    super().__init__()
                    self._app = app
                    self._tps = tps_count
                def write(self, s):
                    super().write(s)
                    s = s.strip()
                    if "Loading config" in s:
                        self._app._set_status("Reading config.xlsx…", self._app.BTN_BLUE)
                    elif "Loading TP" in s:
                        tp = s.split("Loading ")[1].split("…")[0].strip()
                        self._app._set_status(f"Loading {tp} data…", self._app.BTN_BLUE)
                    elif "Building" in s and "report" in s.lower():
                        self._app._set_status(s.lstrip("✅ "), self._app.BTN_BLUE)
                    elif "Exporting Tableau" in s:
                        self._app._set_status("Exporting Tableau CSVs…", self._app.BTN_BLUE)

            self._set_status(f"Loading data from {len(tps)} timepoint(s)…", self.BTN_BLUE)

            log_buffer = StatusCapture(self, len(tps))
            with contextlib.redirect_stdout(log_buffer):
                run_study(folder)

            log = log_buffer.getvalue()

            # Parse outputs from log — match actual print strings from stability_report_v2.py
            # TP reports print as: "  ✅ TP1 report saved: ..."
            # Trends prints as:    "  ✅ Trends report saved: ..."
            # CSV prints as:       "  ✅ Tableau CSVs saved:"
            tp_reports  = [l for l in log.splitlines() if "report saved:" in l and "Trends" not in l]
            trends_done = "Trends report saved" in log
            csv_done    = "Tableau CSVs saved" in log

            self._finish_success(folder, tps, tp_reports,
                                 trends_done, csv_done, log)

        except Exception:
            self._finish_error(
                "Unexpected Error",
                f"An error occurred:\n\n{traceback.format_exc()}"
            )

    # ── UI state helpers ──────────────────────────────────────────────

    def _set_status(self, msg, color=None):
        color = color or self.FG_GREY
        self.root.after(0, lambda: self.status_var.set(msg))
        self.root.after(0, lambda: self.status_lbl.config(fg=color))

    def _finish_success(self, folder, tps, tp_reports,
                        trends_done, csv_done, log):
        def _update():
            self.progress.stop()
            self._set_status("✓ All reports generated successfully.", "#006100")
            self.run_btn.config(state="normal",
                                bg=self.BTN_GREEN, fg=self.BTN_GFONT)

            summary = (
                f"Reports generated successfully!\n\n"
                f"Study folder:      {os.path.basename(folder)}\n"
                f"Timepoints:        {', '.join(tps)}\n"
                f"TP reports:        {len(tp_reports)}\n"
                f"Trends report:     {'✔' if trends_done else '—'}\n"
                f"Tableau CSVs:      {'✔ (3 files)' if csv_done else '—'}\n\n"
                f"All files saved in:\n{folder}"
            )
            messagebox.showinfo("Done!", summary)
        self.root.after(0, _update)

    def _finish_error(self, title, message):
        def _update():
            self.progress.stop()
            self._set_status("Error — see message for details.", self.ERR_RED)
            self.run_btn.config(state="normal",
                                bg=self.BTN_GREEN, fg=self.BTN_GFONT)
            messagebox.showerror(title, message)
        self.root.after(0, _update)


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    app  = StabilityReportApp(root)
    root.mainloop()
