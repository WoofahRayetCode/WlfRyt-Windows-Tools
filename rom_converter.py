#!/usr/bin/env python3
"""
ROM Converter
A GUI tool for bulk converting disc images to CHD format (PS1, PS2, and more)
"""

import os
import subprocess
import shutil
import json
import zipfile
import tarfile
import urllib.request
import urllib.error
from pathlib import Path
from tkinter import Tk, Frame, Label, Button, Entry, Text, Scrollbar, Checkbutton, BooleanVar, filedialog, messagebox, Toplevel
from tkinter import ttk
import threading
import re
import multiprocessing
from concurrent.futures import ThreadPoolExecutor, as_completed
from queue import Queue
import time
import threading as pythread
try:
    import psutil  # For CPU, memory, disk usage
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

# MAME download configuration
MAME_RELEASE_URL = "https://www.mamedev.org/release.html"
MAME_GITHUB_RELEASES_API = "https://api.github.com/repos/mamedev/mame/releases/latest"

# Supported compressed file extensions
COMPRESSED_EXTENSIONS = {'.zip', '.7z', '.rar', '.gz', '.tar', '.tar.gz', '.tgz'}


class ROMConverter:
    def __init__(self, master):
        self.master = master
        master.title("ROM Converter")
        master.geometry("800x600")
        master.resizable(True, True)
        
        # Config file location
        self.config_file = Path.home() / ".rom_converter_config.json"
        
        # Variables
        self.source_dir = ""
        self.delete_originals = BooleanVar(value=False)
        self.move_to_backup = BooleanVar(value=True)
        self.recursive = BooleanVar(value=True)
        self.is_converting = False
        self.cpu_cores = multiprocessing.cpu_count()
        self.log_queue = Queue()
        self.total_original_size = 0
        self.total_chd_size = 0
        self.process_ps1_cues = BooleanVar(value=False)  # Toggle for PS1 CUE processing
        self.process_ps2_isos = BooleanVar(value=False)  # Toggle for PS2 ISO processing
        self.extract_compressed = BooleanVar(value=True)  # Toggle for extracting compressed files
        self.delete_archives_after_extract = BooleanVar(value=False)  # Delete archives after extraction
        self.seven_zip_path = None  # Path to 7z executable for .7z and .rar files
        # Metrics / ETA tracking
        self.total_jobs = 0
        self.completed_jobs = 0
        self.file_start_times = {}
        self.file_durations = []
        self.conversion_start_time = None
        self.initial_disk_write_bytes = 0
        self.last_disk_write_bytes = 0
        self.metrics_running = False
        self.metrics_lock = pythread.Lock()
        self.chdman_path = None  # Will store path to chdman executable
        
        # Load saved configuration
        self.load_config()
        
        # Check for 7-Zip (for .7z and .rar support)
        self.check_7zip()
        
        # Check for chdman
        if not self.check_chdman():
            # Offer to download MAME tools or manual selection
            response = messagebox.askyesnocancel(
                "chdman Not Found",
                "chdman not found!\n\n"
                "Would you like to download MAME tools automatically?\n\n"
                "Yes = Download from mamedev.org\n"
                "No = Manually select chdman.exe\n"
                "Cancel = Exit"
            )
            if response is True:  # Yes - download
                if self.download_mame_tools():
                    if not self.check_chdman():
                        messagebox.showerror("Error", "Failed to find chdman after download")
                        master.destroy()
                        return
                else:
                    master.destroy()
                    return
            elif response is False:  # No - manual selection
                self.browse_chdman()
                if not self.chdman_path:
                    master.destroy()
                    return
            else:  # Cancel
                master.destroy()
                return
        else:
            # chdman found - check for updates
            self.check_for_chdman_update()
        
        self.setup_ui()
    
    def check_chdman(self):
        """Check if chdman is available"""
        # First check for chdman.exe directly next to script
        script_dir = Path(__file__).parent.resolve()
        direct_chdman = script_dir / "chdman.exe"
        if direct_chdman.exists():
            self.chdman_path = str(direct_chdman)
            return True
        
        # Check PATH as fallback
        chdman = shutil.which("chdman")
        if chdman:
            self.chdman_path = chdman
            return True
        
        return False
    
    def get_installed_chdman_version(self):
        """Get the version of the currently installed chdman"""
        if not self.chdman_path:
            return None
        
        try:
            result = subprocess.run(
                [self.chdman_path, '--version'],
                capture_output=True, text=True, timeout=10
            )
            # Output typically like: "chdman - MAME Compressed Hunks of Data (CHD) manager 0.271 (mame0271)"
            # or "chdman - MAME ... 0.283 (mame0283)"
            output = result.stdout + result.stderr
            
            # Look for version pattern like "0.283" or "(mame0283)"
            version_match = re.search(r'\(mame(\d{4})\)', output)
            if version_match:
                return version_match.group(1)
            
            # Alternative: look for version number like "0.283"
            version_match = re.search(r'(\d+\.\d+)', output)
            if version_match:
                ver = version_match.group(1).replace('.', '')
                return ver.zfill(4)
                
        except Exception as e:
            print(f"Error getting chdman version: {e}")
        
        return None
    
    def check_for_chdman_update(self):
        """Check if a newer version of chdman is available"""
        try:
            installed_version = self.get_installed_chdman_version()
            if not installed_version:
                return  # Can't determine installed version, skip update check
            
            latest_version = self.get_latest_mame_version()
            if not latest_version:
                return  # Can't fetch latest version, skip update check
            
            # Compare versions (they're 4-digit strings like "0283")
            if int(latest_version) > int(installed_version):
                response = messagebox.askyesno(
                    "chdman Update Available",
                    f"A newer version of chdman is available!\n\n"
                    f"Installed: MAME {installed_version[0]}.{installed_version[1:]}\n"
                    f"Latest: MAME {latest_version[0]}.{latest_version[1:]}\n\n"
                    "Would you like to update now?"
                )
                if response:
                    self.download_mame_tools()
        except Exception as e:
            print(f"Error checking for chdman update: {e}")
    
    def get_latest_mame_version(self):
        """Fetch the latest MAME version from mamedev.org"""
        try:
            # Parse the release page to get version number
            req = urllib.request.Request(
                MAME_RELEASE_URL,
                headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROM Converter'}
            )
            with urllib.request.urlopen(req, timeout=30) as response:
                html = response.read().decode('utf-8')
            
            # Look for version pattern like "mame0283" or "MAME 0.283"
            version_match = re.search(r'mame(\d{4})', html, re.IGNORECASE)
            if version_match:
                return version_match.group(1)
            
            # Alternative pattern
            version_match = re.search(r'MAME\s+(\d+\.\d+)', html)
            if version_match:
                # Convert 0.283 to 0283
                ver = version_match.group(1).replace('.', '')
                return ver.zfill(4)
            
        except Exception as e:
            print(f"Error fetching MAME version: {e}")
        
        return None
    
    def download_mame_tools(self):
        """Download and extract MAME tools from mamedev.org"""
        try:
            # Get latest version
            version = self.get_latest_mame_version()
            if not version:
                messagebox.showerror("Error", "Could not determine latest MAME version")
                return False
            
            # Construct download URL for Windows x64 binary
            # Format: mame0283b_x64.exe (self-extracting 7z)
            download_url = f"https://github.com/mamedev/mame/releases/download/mame{version}/mame{version}b_64bit.exe"
            
            # Alternative URL patterns to try
            alt_urls = [
                f"https://github.com/mamedev/mame/releases/download/mame{version}/mame{version}b_x64.exe",
                f"https://github.com/mamedev/mame/releases/download/mame{version}/mame{version}b_64bit.exe",
            ]
            
            script_dir = Path(__file__).parent.resolve()
            temp_dir = script_dir / "mame_temp"
            temp_dir.mkdir(exist_ok=True)
            
            download_path = temp_dir / f"mame{version}.exe"
            
            # Show progress dialog
            progress_window = Tk()
            progress_window.title("Downloading MAME Tools")
            progress_window.geometry("400x150")
            progress_window.resizable(False, False)
            
            Label(progress_window, text=f"Downloading MAME {version} tools...", 
                  font=("Arial", 10)).pack(pady=10)
            Label(progress_window, text="This may take a few minutes (file is ~96 MB)", 
                  font=("Arial", 9)).pack()
            
            progress_bar = ttk.Progressbar(progress_window, mode='indeterminate', length=350)
            progress_bar.pack(pady=20)
            progress_bar.start(10)
            
            status_label = Label(progress_window, text="Connecting...")
            status_label.pack()
            
            progress_window.update()
            
            # Try downloading
            downloaded = False
            for url in [download_url] + alt_urls:
                try:
                    status_label.config(text=f"Trying: {url.split('/')[-1]}")
                    progress_window.update()
                    
                    req = urllib.request.Request(
                        url,
                        headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROM Converter'}
                    )
                    
                    with urllib.request.urlopen(req, timeout=300) as response:
                        total_size = int(response.headers.get('content-length', 0))
                        
                        with open(download_path, 'wb') as f:
                            downloaded_size = 0
                            block_size = 1024 * 1024  # 1MB blocks
                            
                            while True:
                                buffer = response.read(block_size)
                                if not buffer:
                                    break
                                f.write(buffer)
                                downloaded_size += len(buffer)
                                
                                if total_size > 0:
                                    percent = (downloaded_size / total_size) * 100
                                    mb_downloaded = downloaded_size / (1024 * 1024)
                                    mb_total = total_size / (1024 * 1024)
                                    status_label.config(text=f"Downloaded: {mb_downloaded:.1f} / {mb_total:.1f} MB ({percent:.1f}%)")
                                else:
                                    mb_downloaded = downloaded_size / (1024 * 1024)
                                    status_label.config(text=f"Downloaded: {mb_downloaded:.1f} MB")
                                progress_window.update()
                    
                    downloaded = True
                    break
                    
                except urllib.error.HTTPError as e:
                    if e.code == 404:
                        continue  # Try next URL
                    raise
                except Exception:
                    continue
            
            if not downloaded:
                progress_window.destroy()
                messagebox.showerror("Error", "Could not download MAME tools from any mirror")
                return False
            
            # Extract using 7-Zip or the self-extracting exe
            status_label.config(text="Extracting chdman.exe...")
            progress_window.update()
            
            # The MAME exe is a self-extracting 7z archive
            # We need 7-Zip to extract just chdman.exe, or run with specific args
            extracted = False
            
            if self.seven_zip_path:
                # Use 7-Zip to extract only chdman.exe directly to script directory
                try:
                    cmd = [
                        self.seven_zip_path, 'e', str(download_path),
                        '-o' + str(script_dir),
                        'chdman.exe',
                        '-y'
                    ]
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
                    if result.returncode == 0 and (script_dir / "chdman.exe").exists():
                        extracted = True
                except Exception as e:
                    print(f"7-Zip extraction error: {e}")
            
            if not extracted:
                # Try running as self-extracting archive with output directory
                try:
                    cmd = [str(download_path), '-o' + str(temp_dir), '-y']
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
                    # Move chdman.exe to script directory
                    temp_chdman = temp_dir / "chdman.exe"
                    if temp_chdman.exists():
                        shutil.move(str(temp_chdman), str(script_dir / "chdman.exe"))
                        extracted = True
                except Exception as e:
                    print(f"Self-extraction error: {e}")
            
            progress_window.destroy()
            
            # Clean up the temp directory and download file
            if temp_dir.exists():
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass
            
            if extracted and (script_dir / "chdman.exe").exists():
                self.chdman_path = str(script_dir / "chdman.exe")
                self.save_config()
                messagebox.showinfo("Success", f"chdman.exe downloaded successfully!\n\nLocation:\n{self.chdman_path}")
                return True
            else:
                messagebox.showerror(
                    "Extraction Failed",
                    "Could not extract chdman.exe from MAME package.\n\n"
                    "Please install 7-Zip and try again, or manually download MAME from:\n"
                    "https://www.mamedev.org/release.html"
                )
                return False
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to download MAME tools:\n{e}")
            return False
    
    def check_7zip(self):
        """Check if 7-Zip is available"""
        # Check common locations on Windows
        common_paths = [
            r"C:\Program Files\7-Zip\7z.exe",
            r"C:\Program Files (x86)\7-Zip\7z.exe",
        ]
        
        # Check PATH first
        seven_zip = shutil.which("7z")
        if seven_zip:
            self.seven_zip_path = seven_zip
            return True
        
        # Check common install locations
        for path in common_paths:
            if os.path.exists(path):
                self.seven_zip_path = path
                return True
        
        return False
    
    def browse_7zip(self):
        """Allow user to manually select 7z executable"""
        filetypes = [("Executable files", "*.exe"), ("All files", "*.*")]
        seven_zip_file = filedialog.askopenfilename(
            title="Select 7z executable",
            filetypes=filetypes
        )
        if seven_zip_file:
            try:
                result = subprocess.run(
                    [seven_zip_file],
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                if "7-zip" in result.stdout.lower() or "7-zip" in result.stderr.lower():
                    self.seven_zip_path = seven_zip_file
                    self.save_config()
                    self.log(f"7-Zip location set to: {seven_zip_file}")
                    if hasattr(self, 'seven_zip_label'):
                        self.seven_zip_label.config(text=self.seven_zip_path or "Not set")
                    messagebox.showinfo("Success", f"7-Zip location set to:\n{seven_zip_file}")
                else:
                    messagebox.showerror("Error", "Selected file does not appear to be 7-Zip")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to verify 7-Zip:\n{e}")
    
    def browse_chdman(self):
        """Allow user to manually select chdman executable"""
        filetypes = [("Executable files", "*.exe"), ("All files", "*.*")]
        chdman_file = filedialog.askopenfilename(
            title="Select chdman executable",
            filetypes=filetypes
        )
        if chdman_file:
            # Verify it's actually chdman by trying to run it
            try:
                result = subprocess.run(
                    [chdman_file, "--help"],
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                if "chdman" in result.stdout.lower() or "chdman" in result.stderr.lower():
                    self.chdman_path = chdman_file
                    self.save_config()
                    self.log(f"chdman location set to: {chdman_file}")
                    # Update UI label if it exists
                    if hasattr(self, 'chdman_label'):
                        self.chdman_label.config(text=self.chdman_path)
                    messagebox.showinfo("Success", f"chdman location set to:\n{chdman_file}")
                else:
                    messagebox.showerror("Error", "Selected file does not appear to be chdman")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to verify chdman:\n{e}")
    
    def save_config(self):
        """Save configuration to JSON file"""
        try:
            config = {
                'source_dir': self.source_dir,
                'delete_originals': self.delete_originals.get(),
                'move_to_backup': self.move_to_backup.get(),
                'recursive': self.recursive.get(),
                'process_ps1_cues': self.process_ps1_cues.get(),
                'process_ps2_isos': self.process_ps2_isos.get(),
                'extract_compressed': self.extract_compressed.get(),
                'delete_archives_after_extract': self.delete_archives_after_extract.get(),
                'chdman_path': self.chdman_path,
                'seven_zip_path': self.seven_zip_path
            }
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=2)
        except Exception as e:
            # Silently fail - don't interrupt user experience
            pass
    
    def load_config(self):
        """Load configuration from JSON file"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                
                # Restore settings
                self.source_dir = config.get('source_dir', '')
                self.delete_originals.set(config.get('delete_originals', False))
                self.move_to_backup.set(config.get('move_to_backup', True))
                self.recursive.set(config.get('recursive', True))
                self.process_ps1_cues.set(config.get('process_ps1_cues', False))
                self.process_ps2_isos.set(config.get('process_ps2_isos', False))
                self.extract_compressed.set(config.get('extract_compressed', True))
                self.delete_archives_after_extract.set(config.get('delete_archives_after_extract', False))
                
                # Restore chdman path if saved and still exists
                saved_chdman = config.get('chdman_path')
                if saved_chdman and os.path.exists(saved_chdman):
                    self.chdman_path = saved_chdman
                
                # Restore 7-Zip path if saved and still exists
                saved_7zip = config.get('seven_zip_path')
                if saved_7zip and os.path.exists(saved_7zip):
                    self.seven_zip_path = saved_7zip
        except Exception as e:
            # Silently fail - use defaults
            pass
    
    def setup_ui(self):
        """Setup the user interface"""
        # Main container
        main_frame = Frame(self.master, padx=10, pady=10)
        main_frame.pack(fill="both", expand=True)
        
        # Directory selection
        dir_frame = Frame(main_frame)
        dir_frame.pack(fill="x", pady=(0, 10))
        
        Label(dir_frame, text="ROM Directory:").pack(side="left", padx=(0, 10))
        
        self.dir_entry = Entry(dir_frame)
        self.dir_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        # Restore saved directory
        if self.source_dir:
            self.dir_entry.insert(0, self.source_dir)
        
        Button(dir_frame, text="Browse", command=self.browse_directory).pack(side="left")
        
        # chdman location
        chdman_frame = Frame(main_frame)
        chdman_frame.pack(fill="x", pady=(0, 10))
        
        Label(chdman_frame, text="chdman:").pack(side="left", padx=(0, 10))
        self.chdman_label = Label(chdman_frame, text=self.chdman_path or "Not set", 
                                  fg="blue", anchor="w")
        self.chdman_label.pack(side="left", fill="x", expand=True, padx=(0, 10))
        Button(chdman_frame, text="Change Location", 
               command=self.browse_chdman).pack(side="left")
        
        # 7-Zip location
        seven_zip_frame = Frame(main_frame)
        seven_zip_frame.pack(fill="x", pady=(0, 10))
        
        Label(seven_zip_frame, text="7-Zip:").pack(side="left", padx=(0, 10))
        self.seven_zip_label = Label(seven_zip_frame, text=self.seven_zip_path or "Not set (optional for .7z/.rar)", 
                                  fg="blue" if self.seven_zip_path else "gray", anchor="w")
        self.seven_zip_label.pack(side="left", fill="x", expand=True, padx=(0, 10))
        Button(seven_zip_frame, text="Set Location", 
               command=self.browse_7zip).pack(side="left")
        
        # Options
        options_frame = Frame(main_frame)
        options_frame.pack(fill="x", pady=(0, 10))
        
        Checkbutton(options_frame, text="Scan subdirectories recursively", 
                   variable=self.recursive).pack(anchor="w")
        
        Checkbutton(options_frame, text="Move original files to backup folder after conversion", 
                   variable=self.move_to_backup, fg="blue").pack(anchor="w")
        
        Checkbutton(options_frame, text="Delete original files after successful conversion", 
                   variable=self.delete_originals, fg="red").pack(anchor="w")

        Checkbutton(options_frame, text="Process PS1 CUE files (.cue)",
                variable=self.process_ps1_cues, fg="green").pack(anchor="w")

        Checkbutton(options_frame, text="Process PS2 ISO files (.iso)",
                variable=self.process_ps2_isos, fg="purple").pack(anchor="w")

        Checkbutton(options_frame, text="Extract compressed files before conversion (.zip, .7z, .rar)",
                variable=self.extract_compressed, fg="orange").pack(anchor="w")

        Checkbutton(options_frame, text="Delete archive files after extraction",
                variable=self.delete_archives_after_extract, fg="red").pack(anchor="w")
        
        # Action buttons
        button_frame = Frame(main_frame)
        button_frame.pack(fill="x", pady=(0, 10))
        
        self.scan_button = Button(button_frame, text="Scan for CUE Files", 
                                 command=self.scan_directory, bg="#4CAF50", fg="white", 
                                 font=("Arial", 10, "bold"))
        self.scan_button.pack(side="left", padx=(0, 10))
        
        self.convert_button = Button(button_frame, text="Start Conversion", 
                                    command=self.start_conversion, bg="#2196F3", fg="white",
                                    font=("Arial", 10, "bold"), state="disabled")
        self.convert_button.pack(side="left", padx=(0, 10))
        
        self.stop_button = Button(button_frame, text="Stop", 
                                 command=self.stop_conversion, bg="#f44336", fg="white",
                                 state="disabled")
        self.stop_button.pack(side="left")
        
        # Second row of buttons for CHD management
        button_frame2 = Frame(main_frame)
        button_frame2.pack(fill="x", pady=(0, 10))
        
        self.move_chd_button = Button(button_frame2, text="Move CHD Files", 
                                     command=self.move_chd_files_dialog, bg="#9C27B0", fg="white",
                                     font=("Arial", 10, "bold"))
        self.move_chd_button.pack(side="left", padx=(0, 10))
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.pack(fill="x", pady=(0, 10))
        
        # Log area
        log_label = Label(main_frame, text="Log Output:", anchor="w")
        log_label.pack(fill="x")
        
        log_frame = Frame(main_frame)
        log_frame.pack(fill="both", expand=True)
        
        scrollbar = Scrollbar(log_frame)
        scrollbar.pack(side="right", fill="y")
        
        self.log_text = Text(log_frame, wrap="word", yscrollcommand=scrollbar.set, 
                            height=15, bg="#f5f5f5")
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.log_text.yview)
        
        # Status bar
        self.status_label = Label(main_frame, text=f"Ready (System: {self.cpu_cores} CPU cores available)", 
                                 anchor="w", relief="sunken")
        self.status_label.pack(fill="x", pady=(10, 0))

        # Metrics label (prominently visible)
        self.metrics_label = Label(main_frame, text="Metrics: idle", anchor="w", 
                                   bg="#e8f4f8", fg="#000000", relief="sunken", 
                                   font=("Arial", 9, "bold"), padx=5, pady=3)
        self.metrics_label.pack(fill="x", pady=(4,0))

        if not PSUTIL_AVAILABLE:
            self.log("psutil not installed. Install with: pip install psutil for resource metrics.")
        
        # Add trace callbacks to save config when options change
        self.delete_originals.trace_add('write', lambda *args: self.save_config())
        self.move_to_backup.trace_add('write', lambda *args: self.save_config())
        self.recursive.trace_add('write', lambda *args: self.save_config())
        self.process_ps1_cues.trace_add('write', lambda *args: self.save_config())
        self.process_ps2_isos.trace_add('write', lambda *args: self.save_config())
        self.extract_compressed.trace_add('write', lambda *args: self.save_config())
        self.delete_archives_after_extract.trace_add('write', lambda *args: self.save_config())
        
        # Start log queue processor
        self.process_log_queue()
    
    def browse_directory(self):
        """Open directory browser"""
        directory = filedialog.askdirectory(title="Select ROM Directory")
        if directory:
            self.source_dir = directory
            self.dir_entry.delete(0, "end")
            self.dir_entry.insert(0, directory)
            self.save_config()
            self.log(f"Selected directory: {directory}")
    
    def log(self, message):
        """Add message to log (thread-safe)"""
        self.log_queue.put(message)
    
    def process_log_queue(self):
        """Process queued log messages from threads"""
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.insert("end", message + "\n")
                self.log_text.see("end")
        except:
            pass
        finally:
            self.master.after(100, self.process_log_queue)
    
    def find_cue_files(self, directory, recursive=True):
        """Find all .cue files in directory"""
        cue_files = []
        path = Path(directory)
        
        if recursive:
            cue_files = list(path.rglob("*.cue"))
        else:
            cue_files = list(path.glob("*.cue"))
        
        return sorted(cue_files)

    def find_compressed_files(self, directory, recursive=True):
        """Find all compressed files in directory"""
        path = Path(directory)
        compressed_files = []
        
        for ext in COMPRESSED_EXTENSIONS:
            if recursive:
                compressed_files.extend(path.rglob(f"*{ext}"))
            else:
                compressed_files.extend(path.glob(f"*{ext}"))
        
        return sorted(compressed_files)
    
    def extract_archive(self, archive_path):
        """Extract a compressed archive to a folder with the same name"""
        archive_path = Path(archive_path)
        
        # Create extraction folder (same name as archive without extension)
        extract_folder = archive_path.parent / archive_path.stem
        
        # Handle multi-extension like .tar.gz
        if archive_path.name.endswith('.tar.gz') or archive_path.name.endswith('.tgz'):
            extract_folder = archive_path.parent / archive_path.name.replace('.tar.gz', '').replace('.tgz', '')
        
        try:
            extract_folder.mkdir(exist_ok=True)
            ext = archive_path.suffix.lower()
            
            # Handle .zip files using Python's built-in zipfile
            if ext == '.zip':
                self.log(f"  📦 Extracting ZIP: {archive_path.name}")
                with zipfile.ZipFile(archive_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_folder)
                self.log(f"  ✅ Extracted to: {extract_folder.name}/")
                return True, extract_folder
            
            # Handle .tar, .tar.gz, .tgz files using Python's tarfile
            elif ext in ['.tar', '.gz', '.tgz'] or archive_path.name.endswith('.tar.gz'):
                self.log(f"  📦 Extracting TAR: {archive_path.name}")
                mode = 'r:gz' if ext in ['.gz', '.tgz'] or archive_path.name.endswith('.tar.gz') else 'r'
                with tarfile.open(archive_path, mode) as tar_ref:
                    tar_ref.extractall(extract_folder)
                self.log(f"  ✅ Extracted to: {extract_folder.name}/")
                return True, extract_folder
            
            # Handle .7z and .rar files using 7-Zip
            elif ext in ['.7z', '.rar']:
                if not self.seven_zip_path:
                    self.log(f"  ⚠️  Cannot extract {ext} file: 7-Zip not configured")
                    return False, None
                
                self.log(f"  📦 Extracting with 7-Zip: {archive_path.name}")
                cmd = [self.seven_zip_path, 'x', str(archive_path), f'-o{extract_folder}', '-y']
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=3600)
                
                if result.returncode == 0:
                    self.log(f"  ✅ Extracted to: {extract_folder.name}/")
                    return True, extract_folder
                else:
                    self.log(f"  ❌ 7-Zip extraction failed: {result.stderr.strip()}")
                    return False, None
            
            else:
                self.log(f"  ⚠️  Unsupported archive format: {ext}")
                return False, None
                
        except zipfile.BadZipFile:
            self.log(f"  ❌ Invalid or corrupted ZIP file: {archive_path.name}")
            return False, None
        except tarfile.TarError as e:
            self.log(f"  ❌ TAR extraction error: {e}")
            return False, None
        except subprocess.TimeoutExpired:
            self.log(f"  ❌ Extraction timeout: {archive_path.name}")
            return False, None
        except Exception as e:
            self.log(f"  ❌ Extraction error: {e}")
            return False, None
    
    def extract_all_archives(self, directory, recursive=True):
        """Find and extract all compressed files in the directory"""
        compressed_files = self.find_compressed_files(directory, recursive)
        
        if not compressed_files:
            self.log("No compressed files found to extract.")
            return []
        
        self.log(f"\n📦 Found {len(compressed_files)} compressed file(s) to extract:")
        for cf in compressed_files:
            self.log(f"   - {cf.name}")
        self.log("")
        
        extracted_folders = []
        for archive in compressed_files:
            success, folder = self.extract_archive(archive)
            if success and folder:
                extracted_folders.append(folder)
                
                # Delete archive if option is enabled
                if self.delete_archives_after_extract.get():
                    try:
                        archive.unlink()
                        self.log(f"  🗑️  Deleted archive: {archive.name}")
                    except Exception as e:
                        self.log(f"  ⚠️  Could not delete archive: {e}")
        
        return extracted_folders

    def find_game_files(self, directory, recursive=True):
        """Find all supported game descriptor files (.cue and optionally .iso)"""
        path = Path(directory)
        files = []
        if recursive:
            if self.process_ps1_cues.get():
                files.extend(path.rglob("*.cue"))
            if self.process_ps2_isos.get():
                files.extend(path.rglob("*.iso"))
        else:
            if self.process_ps1_cues.get():
                files.extend(path.glob("*.cue"))
            if self.process_ps2_isos.get():
                files.extend(path.glob("*.iso"))
        # Sort for stable processing order
        return sorted(files)
    
    def parse_cue_file(self, cue_path):
        """Parse CUE file to find associated BIN files"""
        bin_files = []
        cue_dir = cue_path.parent
        
        try:
            with open(cue_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
                
            # Find FILE entries in CUE
            file_pattern = re.compile(r'FILE\s+"([^"]+)"\s+BINARY', re.IGNORECASE)
            matches = file_pattern.findall(content)
            
            for match in matches:
                bin_path = cue_dir / match
                if bin_path.exists():
                    bin_files.append(bin_path)
                else:
                    self.log(f"  WARNING: Referenced BIN file not found: {match}")
        
        except Exception as e:
            self.log(f"  ERROR parsing CUE file: {e}")
        
        return bin_files
    
    def scan_directory(self):
        """Scan directory for CUE files"""
        if not self.source_dir or not os.path.isdir(self.source_dir):
            messagebox.showwarning("Warning", "Please select a valid directory")
            return
        
        self.log("\n" + "="*60)
        self.log("SCANNING FOR GAME FILES...")
        self.log("="*60)
        
        # First, check for compressed files
        compressed_count = 0
        compressed_size = 0
        if self.extract_compressed.get():
            compressed_files = self.find_compressed_files(self.source_dir, self.recursive.get())
            if compressed_files:
                compressed_count = len(compressed_files)
                self.log(f"\n📦 Found {compressed_count} compressed file(s):")
                for cf in compressed_files:
                    size_mb = cf.stat().st_size / (1024 * 1024)
                    compressed_size += cf.stat().st_size
                    self.log(f"   - {cf.name} ({size_mb:.1f} MB)")
                self.log("\n⚠️  Compressed files will be extracted when you click 'Start Conversion'.")

        game_files = self.find_game_files(self.source_dir, self.recursive.get())

        if not game_files and compressed_count == 0:
            self.log("No game descriptor files found (.cue/.iso) and no compressed files!")
            self.status_label.config(text="No game files found")
            self.convert_button.config(state="disabled")
            return
        
        # If we have compressed files but no game files, still allow conversion
        # (extraction will reveal game files)
        if not game_files and compressed_count > 0:
            self.log(f"\nNo extracted game files found yet, but {compressed_count} compressed file(s) will be extracted.")
            compressed_size_mb = compressed_size / (1024 * 1024)
            self.status_label.config(text=f"Found {compressed_count} compressed file(s) ({compressed_size_mb:.1f} MB) - Ready to extract & convert")
            self.convert_button.config(state="normal")
            return

        ps1_count = 0
        ps2_count = 0
        total_size = 0

        self.log(f"\nFound {len(game_files)} game descriptor file(s):\n")

        for game_file in game_files:
            game_size = 0
            if game_file.suffix.lower() == '.cue':
                ps1_count += 1
                self.log(f"📀 [PS1] {game_file.name}")
                self.log(f"   Path: {game_file}")
                bin_files = self.parse_cue_file(game_file)
                if bin_files:
                    for bin_file in bin_files:
                        size_mb = bin_file.stat().st_size / (1024 * 1024)
                        game_size += bin_file.stat().st_size
                        self.log(f"   └─ {bin_file.name} ({size_mb:.1f} MB)")
                game_size += game_file.stat().st_size  # CUE size (small)
            elif game_file.suffix.lower() == '.iso':
                ps2_count += 1
                iso_size = game_file.stat().st_size
                size_gb = iso_size / (1024 * 1024 * 1024)
                size_mb = iso_size / (1024 * 1024)
                self.log(f"💿 [PS2] {game_file.name}")
                self.log(f"   Path: {game_file}")
                if size_gb >= 1:
                    self.log(f"   └─ ISO size: {size_gb:.2f} GB")
                else:
                    self.log(f"   └─ ISO size: {size_mb:.1f} MB")
                game_size += iso_size
            total_size += game_size
            self.log("")

        total_size_mb = total_size / (1024 * 1024)
        total_size_gb = total_size / (1024 * 1024 * 1024)

        self.log(f"Totals: PS1: {ps1_count}  PS2: {ps2_count}  Combined: {ps1_count + ps2_count}")
        if compressed_count > 0:
            self.log(f"📦 Plus {compressed_count} compressed file(s) to extract")
        if total_size_gb >= 1:
            self.log(f"💾 Current total size: {total_size_gb:.2f} GB ({total_size_mb:.1f} MB)")
        else:
            self.log(f"💾 Current total size: {total_size_mb:.1f} MB")

        status_text = f"Found PS1:{ps1_count} PS2:{ps2_count}"
        if compressed_count > 0:
            status_text += f" + {compressed_count} archives"
        status_text += f" | Size: {total_size_gb:.2f} GB"
        self.status_label.config(text=status_text)
        self.convert_button.config(state="normal")
    
    def convert_to_chd(self, path):
        """Convert a game file (.cue or .iso) to CHD using appropriate chdman mode"""
        chd_path = path.with_suffix('.chd')

        # Skip if CHD already exists
        if chd_path.exists():
            self.log(f"  ⚠️  CHD already exists, skipping: {chd_path.name}")
            return True

        # Determine command and original size basis
        if path.suffix.lower() == '.cue':
            cmd = [self.chdman_path, 'createcd', '-i', str(path), '-o', str(chd_path)]
            original_size = sum(f.stat().st_size for f in self.parse_cue_file(path)) + path.stat().st_size
            label = 'PS1'
        elif path.suffix.lower() == '.iso':
            cmd = [self.chdman_path, 'createdvd', '-i', str(path), '-o', str(chd_path)]
            original_size = path.stat().st_size
            label = 'PS2'
        else:
            self.log(f"  ❌ Unsupported file type: {path.name}")
            return False

        try:
            self.log(f"  Converting ({label}): {path.name} -> {chd_path.name}")
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=1800)

            if result.returncode == 0 and chd_path.exists():
                chd_size = chd_path.stat().st_size
                savings = ((original_size - chd_size) / original_size) * 100 if original_size > 0 else 0

                # Update totals
                self.total_original_size += original_size
                self.total_chd_size += chd_size

                self.log(f"  ✅ Success! Saved {savings:.1f}% space")
                if original_size >= 1024*1024*1024:
                    self.log(f"     Original: {original_size / (1024*1024*1024):.2f} GB -> CHD: {chd_size / (1024*1024*1024):.2f} GB")
                else:
                    self.log(f"     Original: {original_size / (1024*1024):.1f} MB -> CHD: {chd_size / (1024*1024):.1f} MB")
                return True
            else:
                self.log(f"  ❌ Conversion failed: {result.stderr.strip()}")
                return False
        except subprocess.TimeoutExpired:
            self.log(f"  ❌ Timeout: Conversion took too long")
            return False
        except Exception as e:
            self.log(f"  ❌ Exception: {e}")
            return False
    
    def move_to_backup_folder(self, cue_path):
        """Move original CUE and BIN files to backup folder"""
        try:
            # Create backup folder in the same directory as the CUE file
            backup_dir = cue_path.parent / "original_backup"
            backup_dir.mkdir(exist_ok=True)
            
            bin_files = self.parse_cue_file(cue_path)
            
            # Move BIN files
            for bin_file in bin_files:
                if bin_file.exists():
                    dest = backup_dir / bin_file.name
                    # Handle duplicate names
                    counter = 1
                    while dest.exists():
                        dest = backup_dir / f"{bin_file.stem}_{counter}{bin_file.suffix}"
                        counter += 1
                    shutil.move(str(bin_file), str(dest))
                    self.log(f"  📦 Moved to backup: {bin_file.name}")
            
            # Move CUE file
            if cue_path.exists():
                dest = backup_dir / cue_path.name
                counter = 1
                while dest.exists():
                    dest = backup_dir / f"{cue_path.stem}_{counter}{cue_path.suffix}"
                    counter += 1
                shutil.move(str(cue_path), str(dest))
                self.log(f"  📦 Moved to backup: {cue_path.name}")
            
            return True
        except Exception as e:
            self.log(f"  ❌ Error moving files to backup: {e}")
            return False
    
    def delete_original_files(self, cue_path):
        """Delete original CUE and BIN files"""
        try:
            bin_files = self.parse_cue_file(cue_path)
            
            # Delete BIN files
            for bin_file in bin_files:
                if bin_file.exists():
                    bin_file.unlink()
                    self.log(f"  🗑️  Deleted: {bin_file.name}")
            
            # Delete CUE file
            if cue_path.exists():
                cue_path.unlink()
                self.log(f"  🗑️  Deleted: {cue_path.name}")
            
            return True
        except Exception as e:
            self.log(f"  ❌ Error deleting files: {e}")
            return False
    
    def process_single_file(self, cue_file, file_num, total):
        """Process a single CUE file (for parallel execution)"""
        if not self.is_converting:
            return None
        # Record start time for metrics
        with self.metrics_lock:
            self.file_start_times[cue_file] = time.time()
        self.log(f"\n[{file_num}/{total}] Processing: {cue_file.name}")
        
        success = self.convert_to_chd(cue_file)
        
        if success:
            if self.delete_originals.get():
                self.delete_original_files(cue_file)
            elif self.move_to_backup.get():
                self.move_to_backup_folder(cue_file)
        
        return success
    
    def conversion_thread(self):
        """Run conversion in separate thread with parallel processing"""
        
        # First, extract any compressed files if enabled
        if self.extract_compressed.get():
            self.log("\n" + "="*60)
            self.log("EXTRACTING COMPRESSED FILES...")
            self.log("="*60)
            
            extracted_folders = self.extract_all_archives(self.source_dir, self.recursive.get())
            
            if extracted_folders:
                self.log(f"\n✅ Extracted {len(extracted_folders)} archive(s)")
                self.log("Now scanning for game files in extracted folders...\n")
        
        game_files = self.find_game_files(self.source_dir, self.recursive.get())
        total = len(game_files)
        self.total_jobs = total
        self.completed_jobs = 0
        
        if total == 0:
            self.log("No game files to convert!")
            self.is_converting = False
            self.master.after(0, self.conversion_complete)
            return
        
        # Reset size tracking
        self.total_original_size = 0
        self.total_chd_size = 0
        
        self.log("\n" + "="*60)
        self.log("STARTING CONVERSION...")
        self.log(f"Using {self.cpu_cores} CPU cores for parallel processing")
        self.log("="*60 + "\n")
        
        successful = 0
        failed = 0
        completed = 0
        
        # Use ThreadPoolExecutor for parallel processing
        with ThreadPoolExecutor(max_workers=self.cpu_cores) as executor:
            # Submit all conversion jobs
            futures = {executor.submit(self.process_single_file, f, i, total): f for i, f in enumerate(game_files, 1)}
            
            # Process results as they complete
            for future in as_completed(futures):
                if not self.is_converting:
                    self.log("\n⛔ Conversion stopped by user")
                    executor.shutdown(wait=False, cancel_futures=True)
                    break
                
                try:
                    result = future.result()
                    if result is not None:
                        if result:
                            successful += 1
                        else:
                            failed += 1
                        
                        completed += 1
                        # Metrics update
                        with self.metrics_lock:
                            self.completed_jobs = completed
                            started_at = self.file_start_times.get(futures[future])
                            if started_at:
                                self.file_durations.append(time.time() - started_at)
                        progress_value = (completed / total) * 100
                        self.master.after(0, lambda v=progress_value: self.progress.config(value=v))
                
                except Exception as e:
                    failed += 1
                    cue_file = futures[future]
                    self.log(f"❌ Exception processing {cue_file.name}: {e}")
        
        self.log("\n" + "="*60)
        self.log("CONVERSION COMPLETE!")
        self.log("="*60)
        self.log(f"✅ Successful: {successful}")
        self.log(f"❌ Failed: {failed}")
        self.log(f"📊 Total: {total}")
        
        # Display space savings
        if self.total_original_size > 0:
            original_gb = self.total_original_size / (1024 * 1024 * 1024)
            chd_gb = self.total_chd_size / (1024 * 1024 * 1024)
            saved_gb = (self.total_original_size - self.total_chd_size) / (1024 * 1024 * 1024)
            savings_percent = ((self.total_original_size - self.total_chd_size) / self.total_original_size) * 100
            
            self.log("\n" + "-"*60)
            self.log("💾 SPACE SAVINGS:")
            self.log(f"   Original size: {original_gb:.2f} GB")
            self.log(f"   CHD size:      {chd_gb:.2f} GB")
            self.log(f"   Space saved:   {saved_gb:.2f} GB ({savings_percent:.1f}%)")
            self.log("-"*60)
        
        self.is_converting = False
        self.master.after(0, self.conversion_complete)
    
    def start_conversion(self):
        """Start the conversion process"""
        if not self.source_dir or not os.path.isdir(self.source_dir):
            messagebox.showwarning("Warning", "Please select a valid directory")
            return
        
        if self.delete_originals.get() and self.move_to_backup.get():
            messagebox.showwarning(
                "Conflicting Options",
                "Please choose only one option: either move to backup OR delete originals."
            )
            return
        
        if self.delete_originals.get():
            response = messagebox.askyesno(
                "Confirm Deletion",
                "Are you sure you want to delete original files after conversion?\n\n"
                "This action cannot be undone!"
            )
            if not response:
                return
        
        self.is_converting = True
        # Initialize metrics tracking
        self.metrics_running = True
        self.conversion_start_time = time.time()
        self.file_start_times.clear()
        self.file_durations.clear()
        self.total_jobs = len(self.find_game_files(self.source_dir, self.recursive.get()))
        self.completed_jobs = 0
        if PSUTIL_AVAILABLE:
            try:
                io = psutil.disk_io_counters()
                self.initial_disk_write_bytes = io.write_bytes
                self.last_disk_write_bytes = io.write_bytes
            except Exception:
                self.initial_disk_write_bytes = 0
                self.last_disk_write_bytes = 0
        self.convert_button.config(state="disabled")
        self.scan_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.status_label.config(text="Converting...")
        
        # Run conversion in separate thread
        thread = threading.Thread(target=self.conversion_thread, daemon=True)
        thread.start()
        # Start metrics update loop
        self.master.after(500, self.update_metrics)
    
    def stop_conversion(self):
        """Stop the conversion process"""
        self.is_converting = False
        self.stop_button.config(state="disabled")
        self.status_label.config(text="Stopping...")
    
    def conversion_complete(self):
        """Called when conversion is complete"""
        self.convert_button.config(state="normal")
        self.scan_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.progress.config(value=0)
        self.status_label.config(text=f"Ready (System: {self.cpu_cores} CPU cores available)")
        self.metrics_running = False
        self.metrics_label.config(text="Metrics: idle")

    def format_seconds(self, seconds):
        if seconds is None or seconds < 0:
            return "--"
        m, s = divmod(int(seconds), 60)
        h, m = divmod(m, 60)
        if h > 0:
            return f"{h}h {m}m {s}s"
        if m > 0:
            return f"{m}m {s}s"
        return f"{s}s"

    def update_metrics(self):
        if not self.metrics_running:
            return
        with self.metrics_lock:
            completed = self.completed_jobs
            total = self.total_jobs
            durations = list(self.file_durations)
        avg_time = (sum(durations)/len(durations)) if durations else 0
        remaining = max(total - completed, 0)
        overall_eta = avg_time * (remaining / max(self.cpu_cores, 1)) if avg_time else None
        elapsed = time.time() - self.conversion_start_time if self.conversion_start_time else 0
        if PSUTIL_AVAILABLE:
            try:
                cpu = psutil.cpu_percent(interval=None)
                mem = psutil.virtual_memory()
                io = psutil.disk_io_counters()
                written_total = io.write_bytes - self.initial_disk_write_bytes
                rate_write = (io.write_bytes - self.last_disk_write_bytes) / 0.5
                self.last_disk_write_bytes = io.write_bytes
                metrics_text = (
                    f"CPU {cpu:.0f}% | Mem {mem.percent:.0f}% | DiskW {written_total/1024/1024:.1f}MB (+{rate_write/1024/1024:.1f}MB/s) | "
                    f"Jobs {completed}/{total} | Avg {avg_time:.1f}s | Elapsed {self.format_seconds(elapsed)} | ETA {self.format_seconds(overall_eta)}"
                )
            except Exception:
                metrics_text = f"Jobs {completed}/{total} | Avg {avg_time:.1f}s | ETA {self.format_seconds(overall_eta)}"
        else:
            metrics_text = f"Jobs {completed}/{total} | Avg {avg_time:.1f}s | ETA {self.format_seconds(overall_eta)}"
        self.metrics_label.config(text=metrics_text)
        self.status_label.config(text=f"Converting {completed}/{total} ETA {self.format_seconds(overall_eta)}")
        self.master.after(500, self.update_metrics)

    def clean_game_name(self, filename):
        """Remove all parenthetical tags except disc numbers from filename"""
        name = filename
        
        # First, extract disc number if present (to preserve it)
        disc_match = re.search(r'\(Disc\s*\d+\)', name, flags=re.IGNORECASE)
        disc_tag = disc_match.group(0) if disc_match else ""
        
        # Remove ALL parenthetical content (USA, Europe, Rev 1, v1.0, etc.)
        name = re.sub(r'\s*\([^)]+\)', '', name)
        
        # Remove [!] verified dump markers and similar brackets
        name = re.sub(r'\s*\[[^\]]+\]', '', name)
        
        # Clean up any double spaces
        name = re.sub(r'\s+', ' ', name).strip()
        
        # Re-add disc number if it was present
        if disc_tag:
            name = f"{name} {disc_tag}"
        
        return name
    
    def find_chd_files(self, directory, recursive=True):
        """Find all CHD files in directory"""
        path = Path(directory)
        if recursive:
            return sorted(path.rglob("*.chd"))
        else:
            return sorted(path.glob("*.chd"))
    
    def move_chd_files_dialog(self):
        """Open dialog to move CHD files"""
        # Create dialog window
        dialog = Toplevel(self.master)
        dialog.title("Move CHD Files")
        dialog.geometry("600x500")
        dialog.resizable(True, True)
        dialog.transient(self.master)
        dialog.grab_set()
        
        # Source directory
        source_frame = Frame(dialog, padx=10, pady=5)
        source_frame.pack(fill="x")
        
        Label(source_frame, text="Source Folder:").pack(side="left")
        source_entry = Entry(source_frame)
        source_entry.pack(side="left", fill="x", expand=True, padx=5)
        if self.source_dir:
            source_entry.insert(0, self.source_dir)
        
        def browse_source():
            folder = filedialog.askdirectory(title="Select Source Folder")
            if folder:
                source_entry.delete(0, "end")
                source_entry.insert(0, folder)
        
        Button(source_frame, text="Browse", command=browse_source).pack(side="left")
        
        # Destination directory
        dest_frame = Frame(dialog, padx=10, pady=5)
        dest_frame.pack(fill="x")
        
        Label(dest_frame, text="Destination:").pack(side="left")
        dest_entry = Entry(dest_frame)
        dest_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        def browse_dest():
            folder = filedialog.askdirectory(title="Select Destination Folder")
            if folder:
                dest_entry.delete(0, "end")
                dest_entry.insert(0, folder)
        
        Button(dest_frame, text="Browse", command=browse_dest).pack(side="left")
        
        # Options
        options_frame = Frame(dialog, padx=10, pady=5)
        options_frame.pack(fill="x")
        
        remove_locale = BooleanVar(value=True)
        Checkbutton(options_frame, text="Remove locale descriptors from names (USA, Europe, Japan, etc.)", 
                   variable=remove_locale).pack(anchor="w")
        
        recursive_scan = BooleanVar(value=True)
        Checkbutton(options_frame, text="Scan subdirectories", 
                   variable=recursive_scan).pack(anchor="w")
        
        copy_instead = BooleanVar(value=False)
        Checkbutton(options_frame, text="Copy files instead of moving", 
                   variable=copy_instead).pack(anchor="w")
        
        # Scan button and results
        results_frame = Frame(dialog, padx=10, pady=5)
        results_frame.pack(fill="both", expand=True)
        
        # Results listbox with scrollbar
        list_frame = Frame(results_frame)
        list_frame.pack(fill="both", expand=True)
        
        scrollbar = Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        
        results_text = Text(list_frame, wrap="word", yscrollcommand=scrollbar.set, 
                           height=12, bg="#f5f5f5")
        results_text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=results_text.yview)
        
        # Store found files
        found_files = []
        
        def scan_for_chd():
            source = source_entry.get()
            if not source or not os.path.isdir(source):
                messagebox.showwarning("Warning", "Please select a valid source folder")
                return
            
            results_text.delete("1.0", "end")
            found_files.clear()
            
            chd_files = self.find_chd_files(source, recursive_scan.get())
            
            if not chd_files:
                results_text.insert("end", "No CHD files found in the selected folder.\n")
                return
            
            results_text.insert("end", f"Found {len(chd_files)} CHD file(s):\n\n")
            
            total_size = 0
            for chd in chd_files:
                found_files.append(chd)
                size_mb = chd.stat().st_size / (1024 * 1024)
                total_size += chd.stat().st_size
                
                original_name = chd.stem
                if remove_locale.get():
                    clean_name = self.clean_game_name(original_name)
                    if clean_name != original_name:
                        results_text.insert("end", f"📀 {original_name}.chd\n")
                        results_text.insert("end", f"   → {clean_name}.chd ({size_mb:.1f} MB)\n\n")
                    else:
                        results_text.insert("end", f"📀 {original_name}.chd ({size_mb:.1f} MB)\n\n")
                else:
                    results_text.insert("end", f"📀 {original_name}.chd ({size_mb:.1f} MB)\n\n")
            
            total_gb = total_size / (1024 * 1024 * 1024)
            results_text.insert("end", f"\n{'='*50}\n")
            results_text.insert("end", f"Total: {len(chd_files)} files, {total_gb:.2f} GB\n")
        
        def execute_move():
            source = source_entry.get()
            dest = dest_entry.get()
            
            if not source or not os.path.isdir(source):
                messagebox.showwarning("Warning", "Please select a valid source folder")
                return
            
            if not dest:
                messagebox.showwarning("Warning", "Please select a destination folder")
                return
            
            if not found_files:
                messagebox.showwarning("Warning", "Please scan for CHD files first")
                return
            
            # Create destination if it doesn't exist
            dest_path = Path(dest)
            dest_path.mkdir(parents=True, exist_ok=True)
            
            action = "Copying" if copy_instead.get() else "Moving"
            confirm = messagebox.askyesno(
                "Confirm",
                f"{action} {len(found_files)} CHD file(s) to:\n{dest}\n\n"
                f"{'Names will be cleaned (locale removed)' if remove_locale.get() else 'Names unchanged'}\n\n"
                "Continue?"
            )
            
            if not confirm:
                return
            
            results_text.delete("1.0", "end")
            results_text.insert("end", f"{action} files...\n\n")
            
            success_count = 0
            error_count = 0
            
            for chd in found_files:
                try:
                    original_name = chd.stem
                    if remove_locale.get():
                        new_name = self.clean_game_name(original_name) + ".chd"
                    else:
                        new_name = chd.name
                    
                    dest_file = dest_path / new_name
                    
                    # Handle duplicates
                    counter = 1
                    while dest_file.exists():
                        base_name = new_name.rsplit('.', 1)[0]
                        dest_file = dest_path / f"{base_name} ({counter}).chd"
                        counter += 1
                    
                    if copy_instead.get():
                        shutil.copy2(chd, dest_file)
                    else:
                        shutil.move(str(chd), str(dest_file))
                    
                    results_text.insert("end", f"✅ {original_name}.chd → {dest_file.name}\n")
                    success_count += 1
                    
                except Exception as e:
                    results_text.insert("end", f"❌ {chd.name}: {e}\n")
                    error_count += 1
                
                dialog.update()
            
            results_text.insert("end", f"\n{'='*50}\n")
            results_text.insert("end", f"Complete! ✅ {success_count} succeeded, ❌ {error_count} failed\n")
            
            if success_count > 0:
                messagebox.showinfo("Complete", f"Successfully {'copied' if copy_instead.get() else 'moved'} {success_count} file(s)")
        
        # Action buttons
        action_frame = Frame(dialog, padx=10, pady=10)
        action_frame.pack(fill="x")
        
        Button(action_frame, text="Scan for CHD Files", command=scan_for_chd,
               bg="#4CAF50", fg="white", font=("Arial", 10, "bold")).pack(side="left", padx=5)
        
        Button(action_frame, text="Move/Copy Files", command=execute_move,
               bg="#2196F3", fg="white", font=("Arial", 10, "bold")).pack(side="left", padx=5)
        
        Button(action_frame, text="Close", command=dialog.destroy,
               bg="#757575", fg="white", font=("Arial", 10)).pack(side="right", padx=5)


def main():
    # Set multiprocessing start method for Windows to ensure all cores are utilized
    # This prevents issues with the default 'spawn' method on Windows
    if os.name == 'nt':  # Windows
        try:
            multiprocessing.set_start_method('spawn', force=True)
        except RuntimeError:
            # Already set, ignore
            pass
    
    root = Tk()
    app = ROMConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
