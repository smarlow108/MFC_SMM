#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MFCSystem — Bronkhorst Mass Flow Controller live logger (ProPar, safe)
-----------------------------------------------------------------------

✅ Logs flow (kg/h), MFC temperature (°C), and TCsys temperatures (°C).
✅ Auto-saves live Excel every 30 s.
✅ Forces final save on program exit.
✅ Supports labeled samples (e.g. "cathode13", "anode15")
✅ Works with parameter IDs:
       FLOW → 205
       TEMPERATURE → 142
"""

import os
import time
import datetime
import threading
import pandas as pd
import atexit
import propar


class MFCSystem:
    def __init__(self, port="COM6", baud=38400, interval_ms=1000, samples=None, tc_system=None):
        """
        Initialize the Bronkhorst MFC.
        :param port: Serial COM port (e.g., 'COM6')
        :param baud: Baud rate (default 38400)
        :param interval_ms: Logging interval in milliseconds
        :param samples: Optional list like ['cathode12', 'anode12']
        :param tc_system: Optional CAL_TemperatureControllerSystem instance
        """
        self.port = port
        self.baud = baud
        self.interval_ms = interval_ms
        self.samples = samples or ["cathode", "anode"]
        self.tc_system = tc_system  # ✅ optional link to TCsys

        self._is_running = False
        self._thread = None

        # Parameter IDs
        self.FLOW_ID = 205
        self.TEMP_ID = 142

        # Data storage
        self.timestamps = []
        self.datetimes = []
        self.flows = []
        self.temps = []
        self.tc_temps_1 = []
        self.tc_temps_2 = []

        # File setup
        self.base_dir = r"C:\Users\digibat\Desktop\Arksuite-Electrocatalysis\P010\Flow Meter Data"
        os.makedirs(self.base_dir, exist_ok=True)
        s1, s2 = [str(x).replace(" ", "_") for x in self.samples[:2]]
        self._autosave_interval = 30  # seconds
        self._last_autosave = time.time()
        self._live_filename = os.path.join(
            self.base_dir, f"mfc_log_live.xlsx"
        )

        # Auto-save failsafe
        atexit.register(self._force_save_on_exit)

        # Connect to MFC
        try:
            print(f"Connecting to Bronkhorst MFC on {self.port} ({self.baud} baud)...")
            self.device = propar.instrument(self.port, baudrate=self.baud, channel=1)
            print(f"✅ Connected to Bronkhorst MFC ({self.port}, channel 1)")
        except Exception as e:
            print(f"❌ Failed to connect to MFC: {e}")
            self.device = None

    # ---------------------------------------------------------------------
    # MFC Readouts
    # ---------------------------------------------------------------------

    def read_flow(self):
        """Read flow value (kg/h)."""
        if not self.device:
            return None
        try:
            return float(self.device.readParameter(self.FLOW_ID))
        except Exception as e:
            print(f"⚠️ Flow read error: {e}")
            return None

    def read_temperature(self):
        """Read temperature (°C)."""
        if not self.device:
            return None
        try:
            return float(self.device.readParameter(self.TEMP_ID))
        except Exception as e:
            print(f"⚠️ Temperature read error: {e}")
            return None

    # ---------------------------------------------------------------------
    # Logging
    # ---------------------------------------------------------------------

    def start_logging(self):
        """Start background MFC logging."""
        if self._is_running:
            print("⚠️  MFC logging already running.")
            return

        print(f"📈 Starting live MFC logging for {self.samples} (auto-saving)...")
        self._is_running = True
        self.start_time = time.time()

        def log_loop():
            while self._is_running:
                now = datetime.datetime.now()
                timestamp = time.time() - self.start_time
                flow = self.read_flow()
                temp = self.read_temperature()

                # ✅ Read TCsys if available
                tc1 = tc2 = None
                if self.tc_system:
                    try:
                        tc_values = self.tc_system.read_temperature_inputs()
                        if isinstance(tc_values, (list, tuple)) and len(tc_values) >= 2:
                            tc1, tc2 = tc_values[0], tc_values[1]
                    except Exception as e:
                        print(f"⚠️ TCsys read error: {e}")

                self.timestamps.append(timestamp)
                self.datetimes.append(now.strftime("%Y-%m-%d %H:%M:%S"))
                self.flows.append(flow)
                self.temps.append(temp)
                self.tc_temps_1.append(tc1)
                self.tc_temps_2.append(tc2)

                # Live printout
                if tc1 is not None and tc2 is not None:
                    print(f"{now:%H:%M:%S} | {flow:.3f} kg/h | {temp:.2f} °C | TC1={tc1:.2f} °C | TC2={tc2:.2f} °C")
                else:
                    print(f"{now:%H:%M:%S} | {flow:.3f} kg/h | {temp:.2f} °C")

                # Periodic autosave
                if time.time() - self._last_autosave > self._autosave_interval:
                    self._save_live()
                    self._last_autosave = time.time()

                time.sleep(self.interval_ms / 1000.0)

        self._thread = threading.Thread(target=log_loop, daemon=True)
        self._thread.start()

    def stop_mfc_logger(self, run_dir: str = None):
        """
        Stop MFC logging and save final file.
        
        Args:
            run_dir: If provided, saves an additional copy of the final
                    log into this folder (e.g. the E<n>_... run subfolder).
                    The hardcoded base_dir copy is always saved regardless.
        """
        if not self._is_running:
            print("ℹ️  MFC logger not running.")
            return

        print("🛑 Stopping MFC logger...")
        self._is_running = False
        if self._thread:
            self._thread.join(timeout=2.0)
        self._save_final(run_dir=run_dir)

    # ---------------------------------------------------------------------
    # Saving
    # ---------------------------------------------------------------------

    def _create_dataframe(self):
        """Return the log data as a Pandas DataFrame with metadata."""
        df = pd.DataFrame({
            "datetime": self.datetimes,
            "timestamp_s": self.timestamps,
            "flow_kgph": self.flows,
            "temperature_C": self.temps,
            "TC1_C": self.tc_temps_1,
            "TC2_C": self.tc_temps_2,
        })
        df.attrs["Samples"] = f"{self.samples[0]} / {self.samples[1]}"
        df.attrs["Port"] = self.port
        return df

    def _save_live(self):
        """Auto-save partial data to Excel."""
        if not self.timestamps:
            return
        df = self._create_dataframe()
        try:
            with pd.ExcelWriter(self._live_filename, engine="openpyxl") as writer:
                meta = pd.DataFrame({
                    "Metadata": [
                        f"Samples: {self.samples[0]}, {self.samples[1]}",
                        f"Port: {self.port}",
                        f"Autosave time: {datetime.datetime.now()}",
                    ]
                })
                meta.to_excel(writer, index=False, sheet_name="Info")
                df.to_excel(writer, index=False, sheet_name="Data")
            print(f"💾 Auto-saved → {os.path.basename(self._live_filename)}")
        except Exception as e:
            print(f"⚠️ Auto-save failed: {e}")

    def _save_final(self, run_dir: str = None):
        """
        Save full log with timestamped filename.

        Always saves to self.base_dir (hardcoded archive location).
        If run_dir is provided, saves an additional copy there too.
        """
        if not self.timestamps:
            print("⚠️ No MFC data collected.")
            return

        s1, s2    = [str(x).replace(" ", "_") for x in self.samples[:2]]
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename  = f"mfc_log_{s1}_{s2}_{timestamp}.xlsx"
        df        = self._create_dataframe()

        # ── Always save to hardcoded archive directory ────────────────────
        archive_path = os.path.join(self.base_dir, filename)
        try:
            df.to_excel(archive_path, index=False)
            print(f"✅ MFC log saved (archive) → {archive_path}")
        except Exception as e:
            print(f"❌ Failed to save MFC archive log: {e}")

        # ── Also save to run subfolder if provided ────────────────────────
        if run_dir:
            try:
                os.makedirs(run_dir, exist_ok=True)
                run_path = os.path.join(run_dir, filename)
                df.to_excel(run_path, index=False)
                print(f"✅ MFC log saved (run folder) → {run_path}")
            except Exception as e:
                print(f"❌ Failed to save MFC run log: {e}")
 
    def _force_save_on_exit(self):
        """Ensure save if script interrupted — always saves to archive only."""
        if self._is_running:
            print("\n⚠️ Program exiting — forcing MFC save...")
            self._is_running = False
            self._save_final(run_dir=None)   # no run_dir available at exit
