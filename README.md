# FT-710 Parametric EQ Controller

A Windows desktop GUI for reading, editing, and writing the Yaesu **FT-710**'s
two TX parametric EQ profiles (**PRMTRC** / **P PRMTRC**) and the SSB **TX BPF**
width, using **OmniRig** as the CAT transport.

Built-in community presets (rag-chew, broadcast, DX/contest, wide/open) seed
all 3 bands in both profiles, and you can save tweaked settings back as named
user profiles for instant recall.

## Features

- **Read from radio** pulls back everything in one click: both EQ profiles
  (Processor-OFF and Processor-ON), every band's center frequency, level (dB)
  and bandwidth (Q), plus the SSB TX BPF width.
- Edits all 3 bands × 3 parameters × 2 profiles = **18 EQ values** plus the
  global SSB TX BPF width.
- Built-in starter presets:
  - Factory default (flat)
  - Rag-chew (warm)
  - Broadcast (bright) — W0WC-style
  - DX / Contest (punch)
  - Wide / open (as wide as the FT-710's 50–3050 Hz TX BPF allows)
- Save the current editor state as a **named user profile** (persisted to a
  local JSON file) and recall it later.
- Uses OmniRig's `SendCustomCommand` with the `CustomReply` COM event — no
  direct serial port access, so you can share the radio with your logger.

## Requirements

- Windows (the COM/OmniRig stack is Windows-only)
- Python 3.9+
- [OmniRig](http://www.dxatlas.com/OmniRig/) installed and configured for the
  FT-710 on `Rig1` (or `Rig2` — edit `rig_number` in `App.__init__` if needed)
- `pywin32`

Install the Python dependency:

```bash
pip install pywin32
```

## Usage

1. Start OmniRig and confirm the FT-710 rig slot shows `On-line`.
2. Run:

   ```bash
   python ft710_eq_gui.py
   ```

3. Click **Read from radio** to pull **all** current values back from the
   radio in one go — both EQ profiles (all 3 bands × freq/level/bw) and the
   SSB TX BPF width — and populate every editor.
4. Edit values, or pick a preset from the dropdown and click **Load preset →
   editors**.
5. Click **Send to radio** to write the editor state back to the radio.
6. Optionally click **Save as new profile…** to store the current editor state
   as a named user profile.

## CAT reference

From the FT-710 CAT Operation Reference Manual (2211-A / 2306-C):

- TX parametric EQ lives under `EX 03 03 <P3> <value> ;`
  - P3 `02..10` = PRMTRC (Processor-OFF) bands 1/2/3 freq/level/bw
  - P3 `11..19` = P PRMTRC (Processor-ON) bands 1/2/3 freq/level/bw
- SSB TX BPF SEL lives under `EX 01 01 13 <code> ;`
  (`0`=50–3050, `1`=100–2900, `2`=200–2800, `3`=300–2700, `4`=400–2600)

See `EQ_MAP` and `TX_BPF_VALUES` in `ft710_eq_gui.py` for the full mapping.

## Notes

- Both built-in and user-saved presets load all 3 bands in both profiles plus
  the TX BPF width. Nothing is sent to the radio until you click **Send to
  radio**, so loading a preset is safe to experiment with.
- Preset dB/Q values are reasonable starting points — adjust to taste, mic, and
  voice.

## Disclaimer

Not affiliated with Yaesu or with OmniRig's author. Use at your own risk; CAT
commands are written exactly as the radio interprets them and an incorrect
value could leave a parameter in an unexpected state (easily fixed by clicking
**Read from radio** and editing it back).
