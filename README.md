# 🔱 PrinterExorcism

> *The spirits of printers past linger long after they should have departed. Connections to decommissioned hardware, ghost entries left behind by failed GPOs, duplicate copies multiplying silently in the registry — they accumulate, they haunt, and they confuse. PrinterExorcism drives them out.*

A PowerShell module for IT administrators who are done tolerating phantom printers, stale registry connections, and the unholy proliferation of `Copy (2) of Front Desk Main`. It discovers what's lurking, identifies what should be gone, and banishes it — cleanly, thoroughly, and with structured logging so you know exactly what was purged.

---

## What It Does

PrinterExorcism performs a two-phase printer cleanup on Windows machines:

**Phase 1 — User-mode purge.** Runs without elevation. Clears user-space registry keys (`HKCU`) across all mapped printer paths, removes WMI printer objects, and handles the default printer setting if it's pointing at something that no longer exists.

**Phase 2 — Elevated reckoning.** If Phase 1 encounters anything it couldn't remove, it hands off a manifest of the failures to an elevated process and retries. This covers HKLM connection GUIDs, system-level ghost entries, and anything protected from unprivileged removal. A spooler restart follows.

Both phases write to a structured JSON status file and a timestamped log. The orchestrator reads the Phase 1 output, decides if elevation is needed, and handles the UAC prompt — so the caller doesn't have to manage any of that manually.

---

## Capabilities

- **Ghost printer detection and removal** — identifies printers matching patterns like `Copy*`, `PPO*`, and `Front Desk Main*` via both WMI and direct registry enumeration under `HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers`
- **User-space registry cleanup** — removes entries from `Printers\Connections`, `Printers\DevModePerUser`, `Printers\DevModes2`, `Devices`, and `PrinterPorts`
- **HKLM connection purge** — clears accumulated GUID entries from `HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Connections` (requires elevation)
- **GPO-aware cleanup** — with `-CompareGPO`, non-GPO printer connections are removed while policy-deployed connections are preserved
- **Cross-user targeting** — supply `-TargetUser` to target another user's profile; the module mounts their `NTUSER.DAT` hive, performs cleanup, and unmounts cleanly
- **Read-only discovery** — enumerate all printer entries, registry keys, GPO connections, and ghost candidates without making any changes; optionally export to JSON
- **Automated mode** — redirects status files and logs to `%LOCALAPPDATA%\PrinterExorcist` instead of the Desktop, suitable for scripted or RMM-driven deployments
- **Built-in printer preservation** — Microsoft Print to PDF, XPS Document Writer, Fax, OneNote, and Adobe PDF are never touched

---

## Installation

Run this from an elevated PowerShell prompt. The installer will summon the module from GitHub, unpack the sacred arsenal, and drop it into the system module path:

```powershell
irm https://raw.githubusercontent.com/r0tifer/PrinterExorcism/main/Install.ps1 | iex
```

The module installs to `%ProgramFiles%\WindowsPowerShell\Modules\PrinterExorcism` and is immediately importable.

---

## Usage

### Full exorcism (current user)
```powershell
Invoke-PrinterExorcism -FullCleanup -Automated
```

Runs a full two-phase cleanup. Phase 1 executes in user context. If anything fails, Phase 2 launches elevated automatically.

### Target another user's profile
```powershell
Invoke-PrinterExorcism -TargetUser jsmith
```

Mounts `C:\Users\jsmith\NTUSER.DAT`, performs cleanup against that hive, unmounts when finished. Useful for cleaning up a user's printers while they're logged off, or from an admin session.

### GPO-aware cleanup
```powershell
Invoke-PrinterExorcism -CompareGPO
```

Compares active printer connections against the GPO policy key. Connections not covered by policy are removed; GPO-deployed printers are left intact.

### Discover without touching anything
```powershell
Invoke-PrinterExorcism -JSON
```

Runs the discovery module in read-only mode. Enumerates all registry printer keys, WMI printers, GPO connections, default printer, and ghost candidates. Outputs to the terminal and saves a JSON report to `%TEMP%\PrinterDiscovery.<username>.json`.

### Automated / headless deployment
```powershell
Invoke-PrinterExorcism -Automated
```

Suppresses Desktop output. Logs and status files are written to `%LOCALAPPDATA%\PrinterExorcist`. Suitable for RMM scripts, scheduled tasks, or deployment pipelines.

### Combine flags
```powershell
Invoke-PrinterExorcism -TargetUser jsmith -CompareGPO -Automated
```

---

## Use Cases

**Help desk printer reset** — A user's printer list is a graveyard of old mappings from three office moves ago. Run `Invoke-PrinterExorcism -TargetUser <username>` from an admin session to clear the slate while they wait.

**Post-migration cleanup** — After a print server migration or GPO restructure, ghost connections linger in user profiles. Run with `-CompareGPO` to surgically remove everything that isn't covered by current policy.

**Onboarding/offboarding prep** — Ensure a workstation's printer state is clean before reassigning it. Run full cleanup against the outgoing user's profile while they're offboarded.

**RMM-driven remediation** — Deploy with `-Automated` as a remediation script. The JSON status file provides structured output for parsing results or feeding into a reporting workflow.

**"Why does this printer keep coming back?"** — Run with `-JSON` first. The discovery output shows exactly what registry keys exist, which are GPO-linked, and which entries are ghost candidates. Diagnose before you purge.

---

## Output

After each run, two files are written (to Desktop by default, or `%LOCALAPPDATA%\PrinterExorcist` with `-Automated`):

- **`PrinterCleanup.log`** — timestamped log of every action taken, keyed by level (Info, Warning, Critical)
- **`PrinterCleanup.status.json`** — structured summary of the run, including which printers were cleaned, which failed, which phase completed, and whether elevation was used

Phase summary is also printed to the terminal:

```
📦 Phase 1 cleanup summary for: jsmith
   🖨  Printers cleaned:   HP LaserJet 4, OldFrontDesk
   ❌ Printers failed:
   👻 Ghosts detected:    Copy (2) of HP LaserJet 4
   ☠️  Ghosts failed:
```

---

## Requirements

- Windows PowerShell 5.1 or later
- Windows 10 / Windows Server 2016 or later
- Administrator rights are required for Phase 2 (HKLM cleanup and ghost registry removal); Phase 1 runs without elevation

---

## License

[GPL-3.0](LICENSE)
