"""
FT-710 Parametric EQ Controller (OmniRig + Tkinter)
===================================================

Reads back BOTH EQ profiles from the FT-710 (Processor-OFF "PRMTRC" and
Processor-ON "P PRMTRC"), displays all values, lets you edit them, and
writes the changes back via OmniRig's SendCustomCommand.

Built-in community presets (W0WC-style, broadcast, DX punch, wide/open)
seed all 3 bands in both profiles so you can load a full starting point
and tweak from there.

How to run:
    pip install pywin32
    Make sure OmniRig is installed and Rig1 (or Rig2) is configured for
    the FT-710 and shows ONLINE (StatusStr = "On-line").
    python ft710_eq_gui.py

CAT reference: Yaesu FT-710 CAT Operation Reference Manual (2211-A / 2306-C).
Menu path for TX parametric EQ: EX P1=03 (OPERATION SETTING)
                                   P2=03 (TX AUDIO)
                                   P3=02..19 (EQ params, see EQ_MAP below)

OmniRig API note:
    SendCustomCommand(Command, ReplyLength, ReplyEnd)
    - Command is required to be a SAFEARRAY of UI1 (bytes); pass it via
      win32com.client.VARIANT(VT_ARRAY|VT_UI1, b"EX...;").
    - Replies do NOT come back as a return value. They are delivered via
      the CustomReply COM event (dispid 5), which pywin32 hooks as the
      OnCustomReply method on a DispatchWithEvents subclass.
"""

import json
import os
import queue
import threading
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

import pythoncom
import win32com.client


# --------------------------------------------------------------------------
# CAT / EQ definitions
# --------------------------------------------------------------------------

# P3 addresses inside EX 03 03 xx ...
# "off" profile = PRMTRC  (used when speech processor is OFF)
# "on"  profile = P PRMTRC (used when speech processor is ON)
EQ_MAP = {
    "off": {
        1: {"freq": 2,  "level": 3,  "bw": 4},
        2: {"freq": 5,  "level": 6,  "bw": 7},
        3: {"freq": 8,  "level": 9,  "bw": 10},
    },
    "on": {
        1: {"freq": 11, "level": 12, "bw": 13},
        2: {"freq": 14, "level": 15, "bw": 16},
        3: {"freq": 17, "level": 18, "bw": 19},
    },
}

# TX BPF SEL — global per-mode setting (NOT split off/on like the EQ).
# We control the SSB-mode TX BPF because PRMTRC/P PRMTRC are SSB-side EQs.
# CAT: EX P1=01 (RADIO SETTING), P2=01 (MODE SSB), P3=13 (TX BPF SEL)
TX_BPF_P1 = 1
TX_BPF_P2 = 1
TX_BPF_P3 = 13
TX_BPF_VALUES = [
    ("50 - 3050 Hz (widest)",   0),
    ("100 - 2900 Hz",           1),
    ("200 - 2800 Hz (default)", 2),
    ("300 - 2700 Hz",           3),
    ("400 - 2600 Hz (narrow)",  4),
]

# Frequency dropdowns: ("label", numeric_code)
# Band 1: 00=OFF, 01=100Hz .. 07=700Hz (100 Hz steps)
BAND1_FREQS = [("OFF", 0)] + [(f"{100*i} Hz", i) for i in range(1, 8)]
# Band 2: 00=OFF, 01=700Hz .. 09=1500Hz
BAND2_FREQS = [("OFF", 0)] + [(f"{700 + 100*(i-1)} Hz", i) for i in range(1, 10)]
# Band 3: 00=OFF, 01=1500Hz .. 18=3200Hz
BAND3_FREQS = [("OFF", 0)] + [(f"{1500 + 100*(i-1)} Hz", i) for i in range(1, 19)]

BAND_FREQS = {1: BAND1_FREQS, 2: BAND2_FREQS, 3: BAND3_FREQS}

# Popular presets. Each preset has:
#   "bands": {band -> (freq_code, level_db, bw)} applied to bands 2&3 in both
#            off/on profiles (band 1 stays as whatever the radio reported)
#   "tx_bpf": recommended TX BPF code (0..4) — see TX_BPF_VALUES above
# These are reasonable starting points, not dogma — tweak to taste / mic / voice.
PRESETS = {
    "Factory default (flat)": {
        "bands": {1: (0, 0, 0), 2: (0, 0, 0), 3: (0, 0, 0)},
        "tx_bpf": 2,  # 200 - 2800
    },
    # Warmer / fuller rag-chew sound: slight low lift, flat mid, mild high cut
    "Rag-chew (warm)": {
        "bands": {
            1: (3, 3, 5),    # 300 Hz, +3, Q=5
            2: (3, 0, 3),    # 900 Hz, +0, Q=3
            3: (6, -2, 4),   # 2000 Hz, -2, Q=4
        },
        "tx_bpf": 1,  # 100 - 2900
    },
    # W0WC-style "broadcast" TX EQ published for the FT-710
    "Broadcast (bright)": {
        "bands": {
            1: (2, -2, 5),   # 200 Hz,  -2, Q=5
            2: (4, 0, 3),    # 1000 Hz, +0, Q=3
            3: (14, 4, 4),   # 2800 Hz, +4, Q=4
        },
        "tx_bpf": 0,  # 50 - 3050 (widest, for broadcast sound)
    },
    # DX / contest "punch": cut lows, push mids, mild high presence
    "DX / Contest (punch)": {
        "bands": {
            1: (1, -10, 6),  # 100 Hz,  -10, Q=6  (gets rid of thump)
            2: (5, 4, 4),    # 1100 Hz, +4,  Q=4  (intelligibility)
            3: (10, 3, 5),   # 2400 Hz, +3,  Q=5  (presence)
        },
        "tx_bpf": 3,  # 300 - 2700 (tighter, punch through pileup)
    },
    # Wide / open audio within the FT-710's 50-3050 Hz TX BPF: gentle
    # everywhere, broad Qs, a little extra on top for air.
    "Wide / open": {
        "bands": {
            1: (2, 2, 4),    # 200 Hz,  +2, Q=4
            2: (4, 1, 2),    # 1000 Hz, +1, Q=2 (broad)
            3: (16, 3, 3),   # 3000 Hz, +3, Q=3 (broad, airy)
        },
        "tx_bpf": 0,  # 50 - 3050 (widest)
    },
}


USER_PRESETS_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "ft710_user_presets.json"
)


def load_user_presets() -> dict:
    """User presets capture both 'off' and 'on' profiles plus TX BPF.
    Format: {name: {"off": {band: [freq, level, bw]}, "on": {...},
                     "tx_bpf": int}}."""
    if not os.path.exists(USER_PRESETS_FILE):
        return {}
    try:
        with open(USER_PRESETS_FILE, "r") as f:
            raw = json.load(f)
        out = {}
        for name, profiles in raw.items():
            entry = {}
            for pk, val in profiles.items():
                if pk == "tx_bpf":
                    entry["tx_bpf"] = int(val)
                else:
                    entry[pk] = {int(b): tuple(vals) for b, vals in val.items()}
            out[name] = entry
        return out
    except Exception:
        return {}


def save_user_presets(presets: dict) -> None:
    serializable = {}
    for name, profiles in presets.items():
        entry = {}
        for pk, val in profiles.items():
            if pk == "tx_bpf":
                entry["tx_bpf"] = int(val)
            else:
                entry[pk] = {str(b): list(vals) for b, vals in val.items()}
        serializable[name] = entry
    with open(USER_PRESETS_FILE, "w") as f:
        json.dump(serializable, f, indent=2)


def fmt_signed3(val: int) -> str:
    """Yaesu 3-byte signed field: -20..-00 or +00..+10."""
    if val < 0:
        return f"-{abs(val):02d}"
    return f"+{val:02d}"


def parse_signed3(s: str) -> int:
    """Inverse of fmt_signed3. Accepts e.g. '-05', '+03', '+00'."""
    return int(s.replace("+", ""))


def build_ex_set(p1: int, p2: int, p3: int, p4: str) -> bytes:
    return f"EX{p1:02d}{p2:02d}{p3:02d}{p4};".encode("ascii")


def build_ex_read(p1: int, p2: int, p3: int) -> bytes:
    return f"EX{p1:02d}{p2:02d}{p3:02d};".encode("ascii")


# --------------------------------------------------------------------------
# OmniRig wrapper with event handling
# --------------------------------------------------------------------------

class _RigEventSink:
    """
    Attached to OmnirigX via DispatchWithEvents. pywin32 maps the dispinterface
    IOmniRigXEvents methods to On<MethodName> handlers on this class.

    dispid 5 -> CustomReply(RigNumber, Command, Reply)
    """
    reply_queue: queue.Queue = None  # set by OmniRigClient after dispatch

    def OnCustomReply(self, RigNumber, Command, Reply):
        # Command and Reply arrive as tuples of ints (SAFEARRAY of UI1).
        try:
            cmd_bytes = bytes(Command) if Command is not None else b""
            rep_bytes = bytes(Reply) if Reply is not None else b""
        except Exception:
            cmd_bytes = b""
            rep_bytes = b""
        if self.reply_queue is not None:
            self.reply_queue.put((int(RigNumber), cmd_bytes, rep_bytes))

    # Silence the other events (they still fire, we just don't need them).
    def OnVisibleChange(self): pass
    def OnRigTypeChange(self, RigNumber): pass
    def OnStatusChange(self, RigNumber): pass
    def OnParamsChange(self, RigNumber, Params): pass


class OmniRigClient:
    """
    Runs the COM apartment in a dedicated thread so pywin32 events can be
    pumped without fighting the Tk mainloop.
    """

    def __init__(self, rig_number: int = 1):
        self.rig_number = rig_number
        self.reply_queue: queue.Queue = queue.Queue()
        self._cmd_queue: queue.Queue = queue.Queue()
        self._ready = threading.Event()
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._error = None
        self.rig_type = "?"
        self.status_str = "?"

    def start(self, timeout: float = 5.0):
        self._thread.start()
        if not self._ready.wait(timeout):
            raise RuntimeError("OmniRig did not become ready in time")
        if self._error:
            raise self._error

    def stop(self):
        self._stop.set()
        self._cmd_queue.put(None)

    def send_set(self, cmd: bytes):
        self._cmd_queue.put(("set", cmd))

    def send_read(self, cmd: bytes, reply_len: int = 0):
        """reply_len=0 means 'wait for the ';' terminator'."""
        self._cmd_queue.put(("read", cmd, reply_len))

    # -------- worker thread --------
    def _run(self):
        try:
            pythoncom.CoInitialize()
            # DispatchWithEvents needs a real class to graft handlers onto.
            engine = win32com.client.DispatchWithEvents(
                "OmniRig.OmniRigX", _RigEventSink
            )
            # Share our queue with the event sink instance.
            engine.reply_queue = self.reply_queue

            rig = engine.Rig1 if self.rig_number == 1 else engine.Rig2
            try:
                self.rig_type = str(rig.RigType)
                self.status_str = str(rig.StatusStr)
            except Exception:
                pass

            self._ready.set()

            while not self._stop.is_set():
                try:
                    item = self._cmd_queue.get(timeout=0.05)
                except queue.Empty:
                    pythoncom.PumpWaitingMessages()
                    continue

                if item is None:
                    break

                try:
                    kind = item[0]
                    if kind == "set":
                        _, cmd = item
                        self._send(rig, cmd, reply_len=0, wait_term=False)
                    elif kind == "read":
                        _, cmd, rlen = item
                        # For Yaesu, reply is variable-length until ';',
                        # so we use reply_len=0 + terminator ';'.
                        self._send(rig, cmd, reply_len=rlen, wait_term=True)
                except Exception as e:
                    # Put a synthetic "error reply" on the queue so the UI
                    # doesn't deadlock waiting for something that will never
                    # come.
                    self.reply_queue.put((self.rig_number, item[1], f"ERR:{e}".encode()))

                pythoncom.PumpWaitingMessages()

        except Exception as e:
            self._error = e
            self._ready.set()
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    @staticmethod
    def _send(rig, cmd: bytes, reply_len: int, wait_term: bool):
        command = win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_UI1, cmd
        )
        if wait_term:
            reply_end = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_UI1, b";"
            )
        else:
            reply_end = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_UI1, b""
            )
        rig.SendCustomCommand(command, reply_len, reply_end)


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------

class EQBandFrame(ttk.LabelFrame):
    """One band's editor: freq dropdown, level spinbox, bandwidth spinbox."""

    def __init__(self, parent, band: int, title: str):
        super().__init__(parent, text=title, padding=8)
        self.band = band
        self.freqs = BAND_FREQS[band]

        self.freq_var = tk.StringVar(value=self.freqs[0][0])
        self.level_var = tk.IntVar(value=0)
        self.bw_var = tk.IntVar(value=0)

        ttk.Label(self, text="Center freq").grid(row=0, column=0, sticky="w")
        self.freq_cb = ttk.Combobox(
            self, textvariable=self.freq_var, width=10, state="readonly",
            values=[label for label, _ in self.freqs],
        )
        self.freq_cb.grid(row=0, column=1, padx=4, pady=2, sticky="w")

        ttk.Label(self, text="Level (dB)").grid(row=1, column=0, sticky="w")
        self.level_sb = ttk.Spinbox(
            self, from_=-20, to=10, textvariable=self.level_var, width=6,
        )
        self.level_sb.grid(row=1, column=1, padx=4, pady=2, sticky="w")

        ttk.Label(self, text="Bandwidth (Q)").grid(row=2, column=0, sticky="w")
        self.bw_sb = ttk.Spinbox(
            self, from_=0, to=10, textvariable=self.bw_var, width=6,
        )
        self.bw_sb.grid(row=2, column=1, padx=4, pady=2, sticky="w")

    def get(self):
        """Return (freq_code, level_db, bw)."""
        label = self.freq_var.get()
        freq_code = next((c for lbl, c in self.freqs if lbl == label), 0)
        return freq_code, int(self.level_var.get()), int(self.bw_var.get())

    def set(self, freq_code: int, level_db: int, bw: int):
        label = next((lbl for lbl, c in self.freqs if c == freq_code),
                     self.freqs[0][0])
        self.freq_var.set(label)
        self.level_var.set(level_db)
        self.bw_var.set(bw)


class EQProfileFrame(ttk.LabelFrame):
    """A full 3-band EQ profile (either Processor-OFF or Processor-ON)."""

    def __init__(self, parent, title: str):
        super().__init__(parent, text=title, padding=8)
        self.band_frames = {}
        for i, name in enumerate(("Band 1 (low)", "Band 2 (mid)", "Band 3 (high)")):
            band = i + 1
            f = EQBandFrame(self, band, name)
            f.grid(row=0, column=i, padx=4, pady=4, sticky="nsew")
            self.band_frames[band] = f


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("FT-710 Parametric EQ Controller (OmniRig) - M0NXD")
        self.geometry("860x600")

        # --- OmniRig client ---
        self.client = OmniRigClient(rig_number=1)
        try:
            self.client.start(timeout=5)
        except Exception as e:
            messagebox.showerror(
                "OmniRig",
                f"Could not talk to OmniRig:\n\n{e}\n\n"
                "Make sure OmniRig is installed and running, and that Rig1 "
                "is configured for the FT-710.",
            )
            self.destroy()
            return

        # --- Top bar: rig status + connection actions ---
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text=f"Rig: {self.client.rig_type}").pack(side="left")
        ttk.Label(top, text=f"   Status: {self.client.status_str}").pack(side="left")

        # --- Two EQ profile frames stacked ---
        self.off_profile = EQProfileFrame(self, "EQ – Speech Processor OFF (PRMTRC)")
        self.off_profile.pack(fill="x", padx=10, pady=6)

        self.on_profile = EQProfileFrame(self, "EQ – Speech Processor ON (P PRMTRC)")
        self.on_profile.pack(fill="x", padx=10, pady=6)

        # --- TX BPF (global, SSB mode). One setting on the radio, but stored
        # with each profile so loading a profile restores your preferred width.
        bpf_frame = ttk.LabelFrame(self, text="TX BPF (SSB mode)", padding=8)
        bpf_frame.pack(fill="x", padx=10, pady=6)
        ttk.Label(bpf_frame, text="Width:").grid(row=0, column=0, sticky="w")
        self.tx_bpf_var = tk.StringVar(value=TX_BPF_VALUES[2][0])  # 200-2800
        self.tx_bpf_cb = ttk.Combobox(
            bpf_frame, textvariable=self.tx_bpf_var, width=28, state="readonly",
            values=[label for label, _ in TX_BPF_VALUES],
        )
        self.tx_bpf_cb.grid(row=0, column=1, padx=6, sticky="w")

        # --- Bottom action bar ---
        actions = ttk.Frame(self, padding=8)
        actions.pack(fill="x")

        ttk.Button(actions, text="Read from radio",
                   command=self.read_all).pack(side="left", padx=3)
        ttk.Button(actions, text="Send to radio",
                   command=self.send_all).pack(side="left", padx=3)

        self.user_presets = load_user_presets()

        ttk.Label(actions, text="  Preset:").pack(side="left", padx=(20, 2))
        self.preset_var = tk.StringVar(value="Factory default (flat)")
        self.preset_cb = ttk.Combobox(
            actions, textvariable=self.preset_var, width=26, state="readonly",
            values=self._preset_names(),
        )
        self.preset_cb.pack(side="left")
        ttk.Button(actions, text="Load preset → editors",
                   command=self.load_preset).pack(side="left", padx=3)
        ttk.Button(actions, text="Save as new profile…",
                   command=self.save_as_new_profile).pack(side="left", padx=3)

        # --- Status/log area ---
        self.log_var = tk.StringVar(value="Idle.")
        ttk.Label(self, textvariable=self.log_var, anchor="w",
                  padding=6, relief="sunken").pack(fill="x", side="bottom")

        # --- Pending read tracker: maps P3 (int) -> (profile_key, band, kind) ---
        # kind in {"freq","level","bw"}
        self._pending = {}
        self._poll_replies()

        # Fire an initial read so the GUI comes up populated
        self.after(200, self.read_all)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------- Reply handling -------------------------

    def _poll_replies(self):
        """Drain OmniRig reply queue every 50 ms."""
        try:
            while True:
                rig_no, cmd, reply = self.client.reply_queue.get_nowait()
                self._handle_reply(cmd, reply)
        except queue.Empty:
            pass
        self.after(50, self._poll_replies)

    def _handle_reply(self, cmd: bytes, reply: bytes):
        try:
            c = cmd.decode("ascii", errors="replace")
            r = reply.decode("ascii", errors="replace")
        except Exception:
            return

        # Reply format for an EX read: "EX<P1><P2><P3><P4>;"
        if not r.startswith("EX") or not r.endswith(";"):
            self.log_var.set(f"Unexpected reply to {c!r}: {r!r}")
            return
        if len(r) < 9:
            return
        try:
            p1 = int(r[2:4])
            p2 = int(r[4:6])
            p3 = int(r[6:8])
        except ValueError:
            return
        p4 = r[8:-1]  # everything between P3P3 and the final ';'

        # TX BPF SEL reply: EX 01 01 13 <code>
        if (p1, p2, p3) == (TX_BPF_P1, TX_BPF_P2, TX_BPF_P3):
            try:
                self._set_tx_bpf_code(int(p4))
            except Exception as e:
                self.log_var.set(f"Parse error on {r}: {e}")
                return
            self.log_var.set(f"← {r}")
            return

        # Otherwise we only care about EQ (EX 03 03 xx)
        if (p1, p2) != (3, 3):
            return

        # Decide which profile/band/kind this P3 belongs to
        found = None
        for profile_key in ("off", "on"):
            for band in (1, 2, 3):
                addrs = EQ_MAP[profile_key][band]
                for kind, addr in addrs.items():
                    if addr == p3:
                        found = (profile_key, band, kind)
                        break
                if found:
                    break
            if found:
                break
        if not found:
            return

        profile_key, band, kind = found
        target = (self.off_profile if profile_key == "off"
                  else self.on_profile).band_frames[band]
        try:
            if kind == "freq":
                target.set_freq_only(int(p4)) if hasattr(target, "set_freq_only") \
                    else target.set(int(p4), target.get()[1], target.get()[2])
            elif kind == "level":
                fc, _, bw = target.get()
                target.set(fc, parse_signed3(p4), bw)
            elif kind == "bw":
                fc, lvl, _ = target.get()
                target.set(fc, lvl, int(p4))
        except Exception as e:
            self.log_var.set(f"Parse error on {r}: {e}")
            return

        self.log_var.set(f"← {r}")

    # ------------------------- Actions -------------------------

    def _get_tx_bpf_code(self) -> int:
        label = self.tx_bpf_var.get()
        return next((c for lbl, c in TX_BPF_VALUES if lbl == label),
                    TX_BPF_VALUES[2][1])

    def _set_tx_bpf_code(self, code: int) -> None:
        label = next((lbl for lbl, c in TX_BPF_VALUES if c == code),
                     TX_BPF_VALUES[2][0])
        self.tx_bpf_var.set(label)

    def read_all(self):
        """Queue a read for every EQ parameter plus the SSB TX BPF."""
        self.log_var.set("Reading all EQ parameters + TX BPF…")
        for profile_key in ("off", "on"):
            for band in (1, 2, 3):
                for _, p3 in EQ_MAP[profile_key][band].items():
                    self.client.send_read(build_ex_read(3, 3, p3))
        self.client.send_read(build_ex_read(TX_BPF_P1, TX_BPF_P2, TX_BPF_P3))

    def send_all(self):
        """Push every editor value to the radio, including TX BPF."""
        count = 0
        for profile_key, profile in (("off", self.off_profile),
                                      ("on", self.on_profile)):
            for band in (1, 2, 3):
                freq_code, level_db, bw = profile.band_frames[band].get()
                addrs = EQ_MAP[profile_key][band]
                self.client.send_set(build_ex_set(3, 3, addrs["freq"],
                                                  f"{freq_code:02d}"))
                self.client.send_set(build_ex_set(3, 3, addrs["level"],
                                                  fmt_signed3(level_db)))
                self.client.send_set(build_ex_set(3, 3, addrs["bw"],
                                                  f"{bw:02d}"))
                count += 3
        self.client.send_set(build_ex_set(
            TX_BPF_P1, TX_BPF_P2, TX_BPF_P3,
            f"{self._get_tx_bpf_code():d}"))
        count += 1
        self.log_var.set(f"→ Sent {count} EX commands to radio.")

    def _preset_names(self):
        return list(PRESETS.keys()) + sorted(self.user_presets.keys())

    def load_preset(self):
        """Apply the selected preset to the editors.

        Built-in and user-saved presets both load all 3 bands into both
        profiles (Processor-OFF and Processor-ON) plus the TX BPF width.
        """
        name = self.preset_var.get()
        if name in self.user_presets:
            data = self.user_presets[name]
            for profile_key, profile in (("off", self.off_profile),
                                          ("on", self.on_profile)):
                bands = data.get(profile_key, {})
                for band in (1, 2, 3):
                    if band in bands:
                        fc, lvl, bw = bands[band]
                        profile.band_frames[band].set(fc, lvl, bw)
            if "tx_bpf" in data:
                self._set_tx_bpf_code(int(data["tx_bpf"]))
            self.log_var.set(f"Loaded user profile '{name}' (editors only; "
                             "click 'Send to radio' to apply).")
            return

        preset = PRESETS.get(name)
        if not preset:
            return
        bands = preset["bands"]
        for profile in (self.off_profile, self.on_profile):
            for band in (1, 2, 3):
                fc, lvl, bw = bands[band]
                profile.band_frames[band].set(fc, lvl, bw)
        self._set_tx_bpf_code(int(preset["tx_bpf"]))
        self.log_var.set(f"Loaded preset '{name}' into all 3 bands + TX BPF "
                         "(editors only; click 'Send to radio' to apply).")

    def save_as_new_profile(self):
        """Capture the current editor values (which reflect the last read
        from the radio if you just clicked 'Read from radio') as a new
        named user profile. Persists to ft710_user_presets.json."""
        name = simpledialog.askstring(
            "Save profile",
            "Name for this profile:\n\n"
            "Tip: click 'Read from radio' first if you want a fresh\n"
            "snapshot of what the radio currently has.",
            parent=self,
        )
        if not name:
            return
        name = name.strip()
        if not name:
            return
        if name in PRESETS:
            messagebox.showerror(
                "Save profile",
                f"'{name}' is a built-in preset name. Please choose another.",
            )
            return
        if name in self.user_presets:
            if not messagebox.askyesno(
                "Overwrite?",
                f"A profile called '{name}' already exists. Overwrite it?",
            ):
                return

        snapshot = {}
        for profile_key, profile in (("off", self.off_profile),
                                      ("on", self.on_profile)):
            snapshot[profile_key] = {
                band: profile.band_frames[band].get() for band in (1, 2, 3)
            }
        snapshot["tx_bpf"] = self._get_tx_bpf_code()
        self.user_presets[name] = snapshot

        try:
            save_user_presets(self.user_presets)
        except Exception as e:
            messagebox.showerror("Save profile", f"Could not write file:\n{e}")
            return

        self.preset_cb["values"] = self._preset_names()
        self.preset_var.set(name)
        self.log_var.set(f"Saved current editor state as profile '{name}'.")

    # ------------------------- Shutdown -------------------------

    def _on_close(self):
        try:
            self.client.stop()
        except Exception:
            pass
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
