import os
import shutil
import csv
import time
import ctypes
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


# ─────────────────────────────────────────────────────────────────
#  UNC / SHAREPOINT PATH UTILITIES
# ─────────────────────────────────────────────────────────────────

def is_unc_path(path):
    """Returns True for \\server\share style paths."""
    return path.startswith("\\\\") or path.startswith("//")


def is_sharepoint_url(path):
    """Returns True if the path looks like a SharePoint HTTPS URL."""
    p = path.lower()
    return p.startswith("http://") or p.startswith("https://")


def unc_is_accessible(path, timeout=10):
    """
    Check whether a UNC path is reachable.
    Uses a background thread so the GUI does not freeze if unreachable.
    """
    import threading
    result = [False]

    def probe():
        try:
            os.listdir(path)
            result[0] = True
        except Exception:
            result[0] = False

    t = threading.Thread(target=probe, daemon=True)
    t.start()
    t.join(timeout)
    return result[0]


def map_sharepoint_drive(url, log_fn=print):
    """
    Map a SharePoint HTTPS URL to a free drive letter via 'net use'.
    Returns the drive letter string (e.g. 'Z:') on success, else None.
    Only needed when a .lnk target is an HTTPS URL.
    """
    used = set(os.popen("wmic logicaldisk get name").read().split())
    for letter in "ZYXWVUTSRQPONMLKJIHGFEDCB":
        drive = f"{letter}:"
        if drive not in used:
            try:
                result = subprocess.run(
                    ["net", "use", drive, url, "/persistent:no"],
                    capture_output=True, text=True, timeout=30
                )
                if result.returncode == 0:
                    log_fn(f"  🗺️  Mapped {url} → {drive}")
                    return drive
                else:
                    log_fn(f"  ⚠️  net use failed: {result.stderr.strip()}")
            except Exception as e:
                log_fn(f"  ⚠️  Could not map drive: {e}")
            break
    return None


# ─────────────────────────────────────────────────────────────────
#  SHORTCUT (.lnk) RESOLUTION
# ─────────────────────────────────────────────────────────────────

def resolve_lnk(lnk_path, log_fn=print):
    """
    Resolve a Windows .lnk shortcut to a usable folder path.

    Handles three target types that SharePoint shortcuts produce:
      1. Local / already-mapped path  →  use directly
      2. UNC path (\\server\share\…)  →  verify accessible, use directly
      3. HTTPS SharePoint URL         →  attempt 'net use' to map a drive

    Returns a usable directory path string, or None on failure.
    """
    # ── Step 1: read the .lnk with win32com ──
    target = None
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        sc = shell.CreateShortCut(lnk_path)
        target = sc.Targetpath.strip() if sc.Targetpath else None
    except Exception as e:
        log_fn(f"  ⚠️  win32com unavailable, trying fallback LNK reader: {e}")

    # ── Step 2: fallback — parse raw .lnk bytes ──
    if not target:
        try:
            target = _read_lnk_target_raw(lnk_path)
        except Exception:
            pass

    if not target:
        log_fn(f"  ❌  Could not read shortcut: {lnk_path}")
        return None

    log_fn(f"  🔗  Shortcut target: {target}")

    # ── Step 3: resolve by target type ──

    # Case A: local / already-mounted drive letter
    if not is_unc_path(target) and not is_sharepoint_url(target):
        if os.path.isdir(target):
            return target
        log_fn(f"  ⚠️  Local target not found: {target}")
        return None

    # Case B: UNC path  (\\server\share\…)
    if is_unc_path(target):
        log_fn(f"  🌐  UNC path — checking accessibility (VPN required)...")
        if unc_is_accessible(target):
            log_fn(f"  ✅  UNC accessible: {target}")
            return target
        log_fn(f"  ❌  UNC not accessible (VPN not connected?): {target}")
        return None

    # Case C: HTTPS SharePoint URL  (modern OneDrive shortcuts)
    if is_sharepoint_url(target):
        log_fn(f"  🌐  SharePoint URL — attempting drive mapping...")
        drive = map_sharepoint_drive(target, log_fn)
        if drive and os.path.isdir(drive + "\\"):
            return drive + "\\"
        log_fn(f"  ❌  Could not map SharePoint URL: {target}")
        return None

    return None


def _read_lnk_target_raw(lnk_path):
    """
    Minimal raw .lnk parser to extract LocalBasePath or NetworkSharePath.
    Works without win32com.  Raises on failure.
    """
    with open(lnk_path, 'rb') as f:
        data = f.read()

    if data[:4] != b'\x4c\x00\x00\x00':
        raise ValueError("Not a valid .lnk file")

    link_flags = int.from_bytes(data[0x14:0x18], 'little')
    has_id_list  = bool(link_flags & 0x01)
    has_link_info = bool(link_flags & 0x02)

    offset = 0x4C
    if has_id_list:
        id_list_size = int.from_bytes(data[offset:offset+2], 'little')
        offset += 2 + id_list_size

    if not has_link_info:
        raise ValueError("No LinkInfo block")

    li_start = offset
    li_flags = int.from_bytes(data[li_start+8:li_start+12], 'little')
    has_local  = bool(li_flags & 0x01)
    has_net    = bool(li_flags & 0x02)

    local_base_off  = int.from_bytes(data[li_start+16:li_start+20], 'little')
    net_share_off   = int.from_bytes(data[li_start+20:li_start+24], 'little')
    common_path_off = int.from_bytes(data[li_start+24:li_start+28], 'little')

    def read_sz(rel_off):
        abs_off = li_start + rel_off
        end = data.index(b'\x00', abs_off)
        return data[abs_off:end].decode('latin-1', errors='replace')

    if has_local:
        return read_sz(local_base_off) + read_sz(common_path_off)
    if has_net:
        return read_sz(net_share_off) + read_sz(common_path_off)

    raise ValueError("Could not extract path from LinkInfo")


def collect_search_roots(base_folder, log_fn=print):
    """
    Scan base_folder for .lnk shortcuts (up to 3 levels deep, circular-safe).
    Returns a deduplicated list of real folder paths to walk.
    """
    roots = [base_folder]
    seen  = {os.path.normcase(base_folder)}
    _scan_for_shortcuts(base_folder, roots, seen, log_fn, depth=0)
    return roots


def _scan_for_shortcuts(folder, roots, seen, log_fn, depth):
    if depth > 3:
        return
    try:
        for entry in os.scandir(folder):
            if entry.name.lower().endswith('.lnk') and entry.is_file():
                target = resolve_lnk(entry.path, log_fn)
                if target:
                    key = os.path.normcase(target)
                    if key not in seen:
                        seen.add(key)
                        roots.append(target)
                        log_fn(f"  📂  Added search root: {target}")
                        _scan_for_shortcuts(target, roots, seen, log_fn, depth + 1)
    except PermissionError:
        log_fn(f"  ⚠️  Permission denied scanning: {folder}")
    except Exception as e:
        log_fn(f"  ⚠️  Error scanning {folder}: {e}")


# ─────────────────────────────────────────────────────────────────
#  ONEDRIVE HYDRATION (force cloud-only files to download)
# ─────────────────────────────────────────────────────────────────

FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS = 0x00400000
FILE_ATTRIBUTE_RECALL_ON_OPEN        = 0x00040000


def is_cloud_only(path):
    """Returns True if the file is a cloud-only OneDrive placeholder."""
    if is_unc_path(path):
        return False   # Network files are never OneDrive placeholders
    try:
        attrs = ctypes.windll.kernel32.GetFileAttributesW(path)
        if attrs == -1:
            return False
        return bool(attrs & (FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS | FILE_ATTRIBUTE_RECALL_ON_OPEN))
    except Exception:
        return False


def hydrate_file(path, timeout=120):
    """Force OneDrive to download a cloud-only file. Returns True on success."""
    if not is_cloud_only(path):
        return True
    try:
        with open(path, 'rb') as f:
            f.read(1)
    except Exception:
        pass

    deadline = time.time() + timeout
    while time.time() < deadline:
        if not is_cloud_only(path):
            return True
        time.sleep(0.5)
    return False


# ─────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────

def pick_folder(title):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title=title)
    root.destroy()
    return folder


def pick_csv_file():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    filepath = filedialog.askopenfilename(
        title="Select CSV file with filenames",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    root.destroy()
    return filepath


def normalise(filename):
    """Ensure filename has .csv extension."""
    if not filename.lower().endswith('.csv'):
        filename += '.csv'
    return filename


def load_filenames_from_csv(csv_path):
    filenames = []
    with open(csv_path, newline='', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        headers = next(reader, None)

        col_index = 0
        if headers:
            lower_headers = [h.strip().lower() for h in headers]
            for keyword in ['filename', 'file_name', 'file', 'name']:
                if keyword in lower_headers:
                    col_index = lower_headers.index(keyword)
                    break
            else:
                first = headers[0].strip()
                if first:
                    filenames.append(normalise(first))

        for row in reader:
            if row and row[col_index].strip():
                filenames.append(normalise(row[col_index].strip()))

    return filenames


# ─────────────────────────────────────────────────────────────────
#  CORE COPY LOGIC
# ─────────────────────────────────────────────────────────────────

def copy_files(source_folder, destination_folder, files_to_find,
               log_widget=None, progress_var=None):
    os.makedirs(destination_folder, exist_ok=True)

    target_files   = set(files_to_find)
    found_files    = {}   # filename → [path1, path2, ...]
    hydrate_failed = []

    def log(msg):
        print(msg)
        if log_widget:
            log_widget.insert(tk.END, msg + "\n")
            log_widget.see(tk.END)
            log_widget.update()

    # ── Resolve all search roots ──
    log("🔍 Resolving source folder and shortcuts...")
    search_roots = collect_search_roots(source_folder, log_fn=log)
    log(f"\n📂 Search roots ({len(search_roots)} total):")
    for r in search_roots:
        log(f"   • {r}")
    log(f"\n   Files to find: {len(files_to_find)}\n")

    # ── Pass 1: Walk every root, collect all matching paths ──
    for root_path in search_roots:
        log(f"  🔎 Walking: {root_path}")
        try:
            for dirpath, dirs, files in os.walk(root_path):
                for filename in files:
                    if filename in target_files:
                        full_path = os.path.join(dirpath, filename)
                        if filename not in found_files:
                            found_files[filename] = []
                        if full_path not in found_files[filename]:
                            found_files[filename].append(full_path)
        except PermissionError:
            log(f"  ⚠️  Permission denied: {root_path}")
        except Exception as e:
            log(f"  ⚠️  Error walking {root_path}: {e}")

    # Report duplicates (informational — best copy chosen below)
    duplicates = {f: p for f, p in found_files.items() if len(p) > 1}
    if duplicates:
        log(f"\n⚠️  {len(duplicates)} file(s) found in multiple locations:")
        for fname, paths in duplicates.items():
            log(f"   {fname}:")
            for p in paths:
                log(f"     • {p}")

    log(f"\n📋 Matched {len(found_files)} / {len(files_to_find)} file(s).\n")
    log("─" * 60)

    # ── Pass 2: Hydrate + Copy ──
    # Priority: already-local  >  network/UNC  >  cloud-only OneDrive
    copied = 0
    total  = len(found_files)

    for i, (filename, paths) in enumerate(found_files.items(), 1):
        if progress_var:
            progress_var.set(int((i / total) * 100))
            if log_widget:
                log_widget.update()

        local_paths   = [p for p in paths if not is_cloud_only(p) and not is_unc_path(p)]
        network_paths = [p for p in paths if is_unc_path(p)]
        cloud_paths   = [p for p in paths if is_cloud_only(p)]
        ordered       = local_paths + network_paths + cloud_paths

        copied_ok = False
        for source_path in ordered:
            if is_cloud_only(source_path):
                log(f"  ☁️  [{i}/{total}] Downloading from OneDrive: {filename} ...")
                if not hydrate_file(source_path, timeout=120):
                    log(f"  ⏭️  Hydration timed out: {source_path} — trying next...")
                    continue

            dest_path = os.path.join(destination_folder, filename)
            try:
                shutil.copy2(source_path, dest_path)
                copied += 1
                copied_ok = True
                src_label = "network" if is_unc_path(source_path) else "local"
                log(f"  ✅ [{i}/{total}] Copied ({src_label}): {filename}")
                break
            except Exception as e:
                log(f"  ⚠️  Copy failed from {source_path}: {e} — trying next...")

        if not copied_ok:
            log(f"  ❌ [{i}/{total}] All copies failed for: {filename}")
            hydrate_failed.append(filename)

    # ── Summary ──
    not_found = [f for f in files_to_find if f not in found_files]

    if not_found:
        log("\n" + "─" * 60)
        log(f"❌ {len(not_found)} file(s) NOT found anywhere:")
        for f in not_found:
            log(f"   - {f}")

    if hydrate_failed:
        log("\n" + "─" * 60)
        log(f"⏱️  {len(hydrate_failed)} file(s) failed to copy:")
        for f in hydrate_failed:
            log(f"   - {f}")

    log(f"\n{'═' * 60}")
    log(f"  ✅ Copied          : {copied}")
    log(f"  ❌ Not found       : {len(not_found)}")
    log(f"  ⏱️  Failed to copy  : {len(hydrate_failed)}")
    log(f"  ⚠️  Duplicates      : {len(duplicates)}")
    log(f"{'═' * 60}\n")

    if progress_var:
        progress_var.set(100)

    return copied, len(not_found), len(hydrate_failed), len(duplicates)


# ─────────────────────────────────────────────────────────────────
#  GUI
# ─────────────────────────────────────────────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("SharePoint CSV File Copier")
        self.root.geometry("740x720")
        self.root.resizable(True, True)

        self.source_folder  = tk.StringVar()
        self.dest_folder    = tk.StringVar()
        self.csv_path       = tk.StringVar()
        self.file_count_var = tk.StringVar(value="No files loaded")
        self.progress_var   = tk.IntVar(value=0)
        self.files_to_find  = []

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        # ── Info banner ──
        tk.Label(
            self.root,
            text=(
                "ℹ️  Supports: local OneDrive folders · UNC network paths (\\\\server\\share) · "
                "SharePoint HTTPS shortcuts (.lnk).  "
                "Cloud-only files are downloaded automatically.  "
                "VPN must be active for company SharePoint / UNC paths."
            ),
            bg="#d1ecf1", fg="#0c5460", wraplength=700, justify="left", padx=10, pady=8
        ).pack(fill="x", padx=12, pady=(10, 0))

        # ── Folders ──
        folder_frame = ttk.LabelFrame(self.root, text=" 📂  Folders ", padding=10)
        folder_frame.pack(fill="x", **pad)
        folder_frame.columnconfigure(1, weight=1)

        ttk.Label(folder_frame, text="Source Folder:").grid(row=0, column=0, sticky="w")
        ttk.Entry(folder_frame, textvariable=self.source_folder, width=56).grid(
            row=0, column=1, padx=6, sticky="ew")
        ttk.Button(folder_frame, text="Browse", command=self._pick_source).grid(row=0, column=2)

        ttk.Label(folder_frame, text="Destination Folder:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(folder_frame, textvariable=self.dest_folder, width=56).grid(
            row=1, column=1, padx=6, pady=(8, 0), sticky="ew")
        ttk.Button(folder_frame, text="Browse", command=self._pick_dest).grid(row=1, column=2, pady=(8, 0))

        ttk.Label(
            folder_frame,
            text="💡 You can type a UNC path directly, e.g.  \\\\server\\share\\MyFolder",
            foreground="grey"
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))

        # ── File Input ──
        input_frame = ttk.LabelFrame(self.root, text=" 📄  Filenames to Copy ", padding=10)
        input_frame.pack(fill="both", expand=True, **pad)

        self.tab = ttk.Notebook(input_frame)
        self.tab.pack(fill="both", expand=True)

        # CSV Tab
        csv_tab = ttk.Frame(self.tab, padding=10)
        self.tab.add(csv_tab, text="  Upload CSV  ")
        ttk.Label(csv_tab, text="Select a CSV containing the filenames you want to copy:").pack(anchor="w")
        csv_row = ttk.Frame(csv_tab)
        csv_row.pack(fill="x", pady=6)
        ttk.Entry(csv_row, textvariable=self.csv_path, width=52).pack(side="left", fill="x", expand=True)
        ttk.Button(csv_row, text="Browse", command=self._pick_csv).pack(side="left", padx=6)
        ttk.Button(csv_tab, text="✅ Load Filenames", command=self._load_csv).pack(anchor="w", pady=(4, 0))
        ttk.Label(csv_tab, foreground="grey",
                  text=(
                      "\n💡 Tips:\n"
                      "   • Single column — one filename per row (e.g. Q000_111)\n"
                      "   • Multi-column — include a header called 'filename' or 'name'\n"
                      "   • .csv extension added automatically if missing"
                  )).pack(anchor="w")

        # Manual Tab
        manual_tab = ttk.Frame(self.tab, padding=10)
        self.tab.add(manual_tab, text="  Type Filenames  ")
        ttk.Label(manual_tab, text="Enter one filename per line (extension optional):").pack(anchor="w")
        self.manual_text = scrolledtext.ScrolledText(manual_tab, height=6, font=("Courier", 10))
        self.manual_text.pack(fill="both", expand=True, pady=6)
        self.manual_text.insert(tk.END, "Q000_111\nQ000_222\n")
        ttk.Button(manual_tab, text="✅ Load Filenames", command=self._load_manual).pack(anchor="w")

        # File count
        ttk.Label(self.root, textvariable=self.file_count_var, foreground="blue").pack(**pad)

        # ── Progress Bar ──
        pf = ttk.LabelFrame(self.root, text=" ⏳  Progress ", padding=6)
        pf.pack(fill="x", padx=12, pady=(0, 4))
        ttk.Progressbar(pf, variable=self.progress_var, maximum=100, length=680).pack(fill="x")

        # ── Log ──
        lf = ttk.LabelFrame(self.root, text=" 📋  Log ", padding=6)
        lf.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        self.log = scrolledtext.ScrolledText(lf, height=8, font=("Courier", 9))
        self.log.pack(fill="both", expand=True)

        # ── Run Button ──
        ttk.Button(self.root, text="🚀  Start Copying", command=self._run).pack(pady=8)

    def _pick_source(self):
        f = pick_folder("Select SOURCE Folder")
        if f:
            self.source_folder.set(f)

    def _pick_dest(self):
        f = pick_folder("Select DESTINATION Folder")
        if f:
            self.dest_folder.set(f)

    def _pick_csv(self):
        p = pick_csv_file()
        if p:
            self.csv_path.set(p)

    def _load_csv(self):
        path = self.csv_path.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Error", "Please select a valid CSV file first.")
            return
        try:
            self.files_to_find = load_filenames_from_csv(path)
            self.file_count_var.set(f"✅ {len(self.files_to_find)} filename(s) loaded from CSV")
        except Exception as e:
            messagebox.showerror("CSV Error", str(e))

    def _load_manual(self):
        text = self.manual_text.get("1.0", tk.END)
        self.files_to_find = [normalise(l.strip()) for l in text.splitlines() if l.strip()]
        self.file_count_var.set(f"✅ {len(self.files_to_find)} filename(s) loaded")

    def _run(self):
        source = self.source_folder.get().strip()
        dest   = self.dest_folder.get().strip()

        if not source:
            messagebox.showerror("Error", "Please enter or browse to a source folder.")
            return
        if not is_unc_path(source) and not os.path.isdir(source):
            messagebox.showerror("Error", f"Source folder not found:\n{source}")
            return
        if is_unc_path(source) and not unc_is_accessible(source):
            if not messagebox.askyesno(
                "UNC path unreachable",
                f"Cannot reach:\n{source}\n\n"
                "This usually means VPN is not connected or the path is wrong.\n\n"
                "Continue anyway?"
            ):
                return
        if not dest:
            messagebox.showerror("Error", "Please select a destination folder.")
            return
        if not self.files_to_find:
            messagebox.showerror("Error", "No filenames loaded. Please load from CSV or type them in.")
            return

        self.log.delete("1.0", tk.END)
        self.progress_var.set(0)

        copied, missing, timeouts, dupes = copy_files(
            source, dest, self.files_to_find,
            log_widget=self.log,
            progress_var=self.progress_var
        )

        messagebox.showinfo(
            "Done!",
            f"✅ Copied          : {copied} file(s)\n"
            f"❌ Not found       : {missing} file(s)\n"
            f"⏱️  Failed to copy  : {timeouts} file(s)\n"
            f"⚠️  Duplicates      : {dupes} file(s)\n\n"
            f"Saved to:\n{dest}"
        )


# ─────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
