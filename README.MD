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
Start-PrinterExorcismSession
```

### Discovery Mode
```powershell
Start-PrinterExorcismSession -JSON
```

### Full Cleanup (User-Mode First, Then Elevation if Needed)
```powershell
Start-PrinterExorcismSession -FullCleanup
```

### System-Wide Purge (Requires Admin, Affects All Users)
```powershell
Start-PrinterExorcismSession -FullCleanup -Automated -TargetUser "SYSTEM" -CompareGPO
```

### For Demos or Maximum Theatrics
```powershell
Make-PrintersSuffer -FullCleanup -Automated -Verbosity Debug
```

---

## 🎚 Supported Parameters

| Flag             | Description |
|------------------|-------------|
| `-Discovery`     | Discover printers and linked registry keys only |
| `-FullCleanup`   | Perform complete cleanup (registry + printer objects) |
| `-CompareGPO`    | Compare current printers with GPO-deployed ones |
| `-Automated`     | Take all actions without prompting |
| `-RetryOnly`     | Re-run only failed steps from Phase 1 |
| `-TargetUser`    | Perform cleanup on another (offline) user hive |
| `-SystemWide`    | Clean printers across all user hives |
| `-Verbosity`     | Log level: Info, Warning, Critical, Debug |

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
