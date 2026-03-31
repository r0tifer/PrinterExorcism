# PrinterExorcism

🖨️ **PrinterExorcism** is a two-phase PowerShell module that forcefully and safely purges all user-linked and system-linked printers — with surgical precision and a flair for drama. It provides discovery, GPO comparison, user-mode cleanup, elevated retries, and full system-wide printer exorcisms.

## 📦 Features

- ✅ **User-mode Phase 1** cleanup (no admin required)
- 🔐 **Elevated Phase 2** retry for failed steps
- 📜 Structured logging via JSON + CLI logs
- 🔎 Printer discovery (registry, WMI, `Get-Printer`)
- 🧼 Registry cleanup (HKCU + HKLM)
- 🪓 Ghost printer detection and removal
- 📊 GPO comparison mode
- 🧙 `Make-PrintersSuffer` dramatic wrapper included
- 🎛️ Logging levels: `Info`, `Warning`, `Critical`, `Debug`

---

## 🧪 Installation

```powershell
irm https://raw.githubusercontent.com/r0tifer/PrinterExorcism/main/Install.ps1 | iex
```

> **Or clone directly:**
```powershell
git clone https://github.com/r0tifer/PrinterExorcism.git; Import-Module .\PrinterExorcism\PrinterExorcism.psm1
```

---

## 🏁 Usage

### Show CLI Help
```powershell
Invoke-PrinterExorcism
```

### Discovery Mode
```powershell
Invoke-PrinterDiscovery -JSON
```

### Full Cleanup (User-Mode First, Then Elevation if Needed)
```powershell
Invoke-PrinterExorcism -FullCleanup
```

### System-Wide Purge (Requires Admin, Affects All Users)
```powershell
Start-SystemWidePrinterCleanup -Automated -Verbosity Debug
```

### For Demos or Maximum Theatrics
```powershell
Make-PrintersSuffer -FullCleanup -Automated -Verbosity Debug
```

### Advanced Session Entry Point
```powershell
Start-PrinterExorcismSession -FullCleanup
```

`Start-PrinterExorcismSession` remains available for direct session orchestration, including `-JSON` discovery mode.

---

## 🎚 Public Commands

| Command | Purpose |
|---------|---------|
| `Invoke-PrinterExorcism` | Main cleanup entry point for current or target user |
| `Invoke-PrinterDiscovery` | Read-only discovery output for current or target user |
| `Start-PrinterExorcismSession` | Lower-level session orchestration entry point |
| `Start-SystemWidePrinterCleanup` | Full cleanup across all user profiles |
| `Make-PrintersSuffer` | Dramatic wrapper around the main cleanup flow |

---

## 📁 Project Structure

```
PrinterExorcism
├── PrinterExorcism.psm1
├── PrinterExorcism.psd1
├── Exorcist
│   ├── Common.ps1
│   └── Private
│       ├── PrinterExorcist.ps1
│       ├── Watch-PrinterExorcism.ps1
│       └── Discover-Printers.ps1
```

---

## 🧾 License

This project is licensed under the **GNU General Public License v3.0**. See the [LICENSE](https://github.com/r0tifer/PrinterExorcism/blob/main/LICENSE) file for details.

---

## 👨‍🔧 Author

Created with righteous indignation by **Michael Levesque**.
